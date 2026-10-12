/**
 * S10.2 real-browser performance receipt helpers: present detection, input to
 * present pairing, sample summaries, importer throughput windows, threshold
 * evaluation and the hash-tagged receipt. Pure functions; no wall clock.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  classifyPresent, evaluateReceipt, finalizeReceipt, importThroughput, pairInputsToPresents, summarize, verifyReceipt,
  type ReceiptDraft,
} from '../src/debug/harness/viewerPerf'

const profile = JSON.parse(readFileSync(join(process.cwd(), 'tests/fixtures/xlsx-viewer/performance-profile.json'), 'utf8'))

describe('classifyPresent', () => {
  const zero = { basePaints: 0, tileHits: 0, prefetched: 0, overlayPaints: 0 }
  it('counts visible tile work as a present and ignores overlay-only or prefetch-only reports', () => {
    expect(classifyPresent(null, { ...zero, tileHits: 6 })).toBe(true)
    expect(classifyPresent(zero, { ...zero, basePaints: 4 })).toBe(true)
    expect(classifyPresent({ ...zero, tileHits: 6 }, { ...zero, tileHits: 6, overlayPaints: 1 })).toBe(false)
    expect(classifyPresent({ ...zero, basePaints: 4 }, { ...zero, basePaints: 6, prefetched: 2 })).toBe(false)
    expect(classifyPresent({ ...zero, basePaints: 4 }, { ...zero, basePaints: 7, prefetched: 2 })).toBe(true)
    expect(classifyPresent(null, zero)).toBe(false)
  })
})

describe('pairInputsToPresents', () => {
  it('pairs each input with the first present at or after it and counts coalesced inputs', () => {
    const { latencies, coalesced, unmatched } = pairInputsToPresents([0, 16, 20, 40], [5, 30, 31])
    expect(latencies).toEqual([5, 14, 10])
    expect(coalesced).toBe(1) // 16 and 20 share the present at 30
    expect(unmatched).toBe(1) // 40 has no present after it
  })
})

describe('summarize', () => {
  it('reports nearest-rank p95 only at or above the minimum sample count', () => {
    const values = Array.from({ length: 300 }, (_, i) => i + 1)
    expect(summarize(values, 300)).toMatchObject({ status: 'ok', samples: 300, p50: 150, p95: 285, max: 300, min: 1 })
    const few = summarize([3, 1, 2], 30) as { status: string; p95: number | null; p50: number }
    expect(few.status).toBe('insufficient-samples')
    expect(few.p95).toBeNull()
    expect(few.p50).toBe(2)
    expect(summarize([], 30).status).toBe('no-samples')
  })
})

describe('importThroughput', () => {
  it('measures post-inventory throughput from the inventory event to done, plus a mid-window sheet slope', () => {
    const MiB = 1024 * 1024
    const events = [
      { t: 100, phase: 'inventory' as const, expandedBytes: 0, populatedCells: 0, compressedRead: 0 },
      { t: 150, phase: 'sharedStrings' as const, expandedBytes: 1 * MiB, populatedCells: 0, compressedRead: 0 },
      { t: 200, phase: 'sheet' as const, expandedBytes: 2 * MiB, populatedCells: 100_000, compressedRead: 0 },
      { t: 700, phase: 'sheet' as const, expandedBytes: 12 * MiB, populatedCells: 600_000, compressedRead: 0 },
      { t: 1200, phase: 'sheet' as const, expandedBytes: 22 * MiB, populatedCells: 1_100_000, compressedRead: 0 },
      { t: 1200, phase: 'done' as const, expandedBytes: 22 * MiB, populatedCells: 1_100_000, compressedRead: 0 },
    ]
    const result = importThroughput(events)!
    expect(result.windowMs).toBe(1100)
    expect(result.expandedMiBPerSec).toBeCloseTo(22 / 1.1, 6)
    expect(result.cellsPerSec).toBeCloseTo(1_100_000 / 1.1, 3)
    expect(result.sheetSlopeMiBPerSec).toBeCloseTo(20, 6) // (22-2) MiB over 1.0 s of sheet pages
    expect(importThroughput([{ t: 0, phase: 'inventory', expandedBytes: 0, populatedCells: 0, compressedRead: 0 }])).toBeNull()
  })
})

function draft(overrides: Partial<ReceiptDraft['measurements']> = {}): ReceiptDraft {
  const frame = (p95: number) => ({ status: 'ok' as const, samples: 300, p50: p95 / 2, p95, max: p95 + 1, min: 1, mean: p95 / 2 })
  const op = (p95: number) => ({ status: 'ok' as const, samples: 30, p50: p95 / 2, p95, max: p95 + 1, min: 1, mean: p95 / 2 })
  return {
    schema: 'officeview-viewer-perf/1',
    mode: 'release',
    environment: { commit: 'abc', dirty: false, browser: 'chrome', browserVersion: '155', channel: 'chrome', headless: true, profile: 'desktop', viewport: { width: 1440, height: 900 }, dpr: 2, machine: 'M4 Pro', os: 'macOS', loadAverage: [1, 1, 1], contention: 'note', startedAt: 'a', finishedAt: 'b' },
    fixtures: [],
    measurements: {
      scroll: { ordinary: { inputToPaint: frame(10), frameInterval: frame(16.7) }, wrapped: { inputToPaint: frame(12), frameInterval: frame(16.7) }, styled: { inputToPaint: frame(9), frameInterval: frame(16.7) } },
      ack: { sort: op(20), filter: op(20), find: op(20), cancel: op(20) },
      firstViewport: { F1: op(400), F2: { ...op(900), status: 'insufficient-samples', p95: null, samples: 10 }, F3: op(800) },
      import: { F1: { expandedMiBPerSec: 30, cellsPerSec: 200_000 }, F2: { expandedMiBPerSec: 25, cellsPerSec: 150_000 }, F3: { expandedMiBPerSec: 21, cellsPerSec: 120_000 } },
      memory: { mainOwnedBytes: 10 * 1024 * 1024, workerOwnedBytes: 300 * 1024 * 1024, bitmapPeakBytes: 100 * 1024 * 1024 },
      ...overrides,
    },
    raw: {},
  }
}

describe('evaluateReceipt', () => {
  it('passes desktop rows within the frozen ceilings and floors, reuses kernel evidence and never claims mobile', () => {
    const rows = evaluateReceipt(draft(), profile)
    const byId = Object.fromEntries(rows.filter(r => r.profile === 'desktop').map(r => [r.id, r]))
    expect(byId['warm-viewport-frame'].outcome).toBe('pass')
    expect(byId['input-acknowledgment'].outcome).toBe('pass')
    expect(byId['numeric-sort-1m'].outcome).toBe('pass')
    expect(byId['numeric-sort-1m'].source).toMatch(/S8\.6/)
    expect(byId['import-expanded-throughput'].outcome).toBe('pass')
    expect(byId['main-owned-memory'].outcome).toBe('pass')
    // F2 first viewport ran with 10 cold samples: p95 refused, so the row is a
    // limitation (max against the ceiling), never a silent p95 pass.
    expect(byId['progressive-first-viewport'].outcome).toBe('pass-limited')
    for (const row of rows.filter(r => r.profile === 'mobile')) expect(row.outcome).toBe('not-run')
  })

  it('misses on a slow frame series, a low throughput floor or an over-budget pool, and is not-run when missing', () => {
    const slow = draft({ scroll: { ordinary: { inputToPaint: { status: 'ok', samples: 300, p50: 9, p95: 21, max: 40, min: 1, mean: 10 }, frameInterval: { status: 'ok', samples: 300, p50: 16.7, p95: 33.3, max: 50, min: 16, mean: 18 } } } })
    const rows = evaluateReceipt(slow, profile).filter(r => r.profile === 'desktop')
    const byId = Object.fromEntries(rows.map(r => [r.id, r]))
    expect(byId['warm-viewport-frame'].outcome).toBe('miss')
    const low = evaluateReceipt(draft({ import: { F1: { expandedMiBPerSec: 5, cellsPerSec: 200_000 } }, memory: { mainOwnedBytes: 80 * 1024 * 1024, workerOwnedBytes: null, bitmapPeakBytes: 100 * 1024 * 1024 } }), profile)
    const lowById = Object.fromEntries(low.filter(r => r.profile === 'desktop').map(r => [r.id, r]))
    expect(lowById['import-expanded-throughput'].outcome).toBe('miss')
    expect(lowById['main-owned-memory'].outcome).toBe('miss')
    expect(lowById['worker-owned-memory'].outcome).toBe('not-run')
  })
})

describe('receipt hash', () => {
  it('detects any modification of a finalized receipt', async () => {
    const receipt = await finalizeReceipt(draft(), profile)
    expect(await verifyReceipt(receipt)).toEqual({ ok: true })
    const tampered = { ...receipt, measurements: { ...receipt.measurements, memory: { ...receipt.measurements.memory!, mainOwnedBytes: 1 } } }
    expect((await verifyReceipt(tampered)).ok).toBe(false)
  })
})

describe('retained S10.2 Chrome receipt', () => {
  it('verifies, is cited by the frozen profile, and agrees with it row by row without touching frozen values', async () => {
    const evidence = profile.viewerEvidence as { results: string; hash: string; commit: string; desktopAdmitted: boolean; rows: Record<string, { outcome: string; value: number | null }> }
    const receipt = JSON.parse(readFileSync(join(process.cwd(), evidence.results), 'utf8'))
    expect(await verifyReceipt(receipt)).toEqual({ ok: true })
    expect(receipt.hash).toBe(evidence.hash)
    expect(receipt.environment.commit).toBe(evidence.commit)
    expect(receipt.environment.dirty).toBe(false)
    expect(receipt.desktopAdmitted).toBe(evidence.desktopAdmitted)
    // The evaluation recomputed from the receipt's own measurements matches what it recorded.
    expect(evaluateReceipt(receipt, profile)).toEqual(receipt.evaluation)
    for (const [id, row] of Object.entries(evidence.rows)) {
      const recorded = receipt.evaluation.find((r: { id: string; profile: string }) => r.id === id && r.profile === 'desktop')
      expect(recorded.outcome, id).toBe(row.outcome)
      expect(recorded.value, id).toBe(row.value)
      const threshold = profile.thresholds.find((t: { id: string }) => t.id === id)
      if (threshold) {
        expect(threshold.desktop.frozenMs, id).toBe(threshold.desktop.specMs)
        expect(threshold.desktop.outcome, id).toBe(row.outcome === 'miss' ? 'miss' : 'pass')
      }
    }
    for (const row of receipt.evaluation.filter((r: { profile: string }) => r.profile === 'mobile')) expect(row.outcome).toBe('not-run')
  })
})
