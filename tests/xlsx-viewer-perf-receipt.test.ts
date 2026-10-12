/**
 * S10.2 real-browser performance receipt helpers: present detection, input to
 * present pairing, sample summaries, importer throughput windows, threshold
 * evaluation and the hash-tagged receipt. Pure functions; no wall clock.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  ACK_CENSORED_MS, RELEASE_IDENTITY, classifyPresent, finalizeReceipt, importThroughput, pairInputsToPresents, summarize, validateReceipt, verifyReceipt,
  type EvidenceContext, type ReceiptDraft, type ViewerReceipt,
} from '../src/debug/harness/viewerPerf'

const profile = JSON.parse(readFileSync(join(process.cwd(), 'tests/fixtures/xlsx-viewer/performance-profile.json'), 'utf8'))
const manifest = JSON.parse(readFileSync(join(process.cwd(), 'tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json'), 'utf8'))
const ctx: EvidenceContext = { profile, fixtures: manifest.scales.full }

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


// ---- a synthetic, fully valid raw record set --------------------------------
const MiB = 1024 * 1024
function timeline(frames: number, workMs: number) {
  const f = Array.from({ length: frames + 1 }, (_, i) => i * 16.67)
  const events = f.slice(0, frames).map(t => t + 0.5)
  return { frames: f, steps: events, events, presents: events.map(t => t + workMs) }
}
const ackRecords = (n: number, ms: number) => Array.from({ length: n }, (_, i) => ({ inputAt: 1000 + i * 100, inputType: 'pointerdown', mutationAt: 1000 + i * 100 + ms - 1, presentAt: 1000 + i * 100 + ms, what: 'button @aria-busy', statusBefore: '', statusAfter: '' }))
const runs = (n: number, ms: number) => ({ runs: Array.from({ length: n }, () => ({ firstPresent: ms - 50, firstUsable: ms, status: ms - 60, complete: ms * 4, usableStatus: 'Visible body 0' })) })
const importRun = () => ({ outcome: 'complete', events: [
  { t: 0, phase: 'inventory', expandedBytes: 0, populatedCells: 0, compressedRead: 0 },
  { t: 100, phase: 'sheet', expandedBytes: 10 * MiB, populatedCells: 200_000, compressedRead: 0 },
  { t: 500, phase: 'done', expandedBytes: 30 * MiB, populatedCells: 1_000_000, compressedRead: 0 },
] })
const checkpoint = (heapBytes: number, workerPeak: number) => ({ heapBytes, worker: { kind: 'stats', sessions: [{ memory: { peakBytes: workerPeak } }] }, uasm: { available: true, worker: workerPeak, canvas: 40 * MiB }, frame: { peakBitmapBytes: 100 * MiB } })
const memorySession = () => ({ checkpoints: { baseline: { heapBytes: 5 * MiB }, loaded: checkpoint(20 * MiB, 200 * MiB), afterOps: checkpoint(25 * MiB, 220 * MiB), closed: { heapBytes: 6 * MiB } }, polls: [{ workerPeak: 210 * MiB }], timeouts: [], liveHeapSamples: [12 * MiB, 30 * MiB, 18 * MiB] })
function validRaw(): Record<string, unknown> {
  return {
    scroll: Object.fromEntries(['ordinary', 'wrapped', 'styled'].map(k => [k, { importComplete: true, timeline: timeline(300, 3) }])),
    ack: { sort: { records: ackRecords(30, 17) }, filterApply: { records: ackRecords(10, 33) }, find: { records: ackRecords(30, 14) }, cancel: { records: ackRecords(30, 16) } },
    firstViewport: { F1: runs(30, 480), F2: runs(10, 600), F3: runs(30, 670) },
    import: Object.fromEntries(['F1', 'F2', 'F3'].map(k => [k, { runs: [importRun(), importRun(), importRun()] }])),
    memory: { F2: memorySession(), F4: memorySession() },
    errors: [],
  }
}
function validDraft(): ReceiptDraft {
  return {
    schema: 'officeview-viewer-perf/1',
    mode: 'release',
    environment: { commit: 'a'.repeat(40), dirty: false, browser: 'chrome', browserVersion: '155.0.8059.39', channel: 'chrome', headless: true, profile: 'desktop', viewport: { width: 1440, height: 900 }, dpr: 2, machine: 'M4 Pro', os: 'macOS', loadAverage: [1, 1, 1], contention: 'note', startedAt: 'a', finishedAt: 'b', build: RELEASE_IDENTITY.build },
    fixtures: ['F1', 'F2', 'F3', 'F4', 'wrapped', 'styled'].map(id => ({ id, file: `${id}.xlsx`, bytes: 1, sha256: 'x', scale: id, ...manifest.scales.full[id] })),
    raw: validRaw(),
  }
}
/** A main-owned ledger does not exist yet; tests of the rest of admission pretend one did. */
const ledgerWorld = (r: ViewerReceipt) => r.evidenceProblems.filter(p => !p.startsWith('main-owned-memory'))

describe('evaluation from raw records', () => {
  it('derives every row from raw records; with valid evidence only the main-memory row stays pending', async () => {
    const receipt = await finalizeReceipt(validDraft(), ctx)
    const byId = Object.fromEntries(receipt.evaluation.filter(r => r.profile === 'desktop').map(r => [r.id, r]))
    expect(byId['warm-viewport-frame'].outcome).toBe('pass')
    expect(byId['warm-viewport-frame'].value).toBeCloseTo(16.67, 2)
    expect(byId['input-acknowledgment'].outcome).toBe('pass-limited') // filter apply: documented 10-sample exception
    expect(byId['progressive-first-viewport'].outcome).toBe('pass-limited') // F2 cold opens: documented exception
    expect(byId['numeric-sort-1m'].source).toMatch(/S8\.6/)
    expect(byId['import-expanded-throughput'].outcome).toBe('pass')
    expect(byId['worker-owned-memory'].value).toBe(220 * MiB)
    expect(byId['main-owned-memory'].outcome).toBe('pending')
    expect(byId['main-owned-memory'].value).toBe(25 * MiB) // sampled live peak 30 - baseline 5
    expect(ledgerWorld(receipt)).toEqual([])
    expect(receipt.desktopAdmitted).toBe(false) // pending main peak blocks admission
    for (const row of receipt.evaluation.filter(r => r.profile === 'mobile')) expect(row.outcome).toBe('not-run')
    expect((await validateReceipt(receipt, ctx)).problems).toEqual(['main-owned-memory: no complete main-owned application ledger; the live peak is not demonstrated (row pending)'])
  })

  it('a sampled live peak above the budget is a demonstrated miss', async () => {
    const d = validDraft()
    ;(d.raw.memory as { F2: { liveHeapSamples: number[] } }).F2.liveHeapSamples = [100 * MiB]
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evaluation.find(r => r.id === 'main-owned-memory' && r.profile === 'desktop')!.outcome).toBe('miss')
  })
})

describe('adversarial receipts (review finding 1)', () => {
  it('(a) smoke mode, dirty WebKit identity, bad commit, no fixtures and no raw data never admit', async () => {
    const d = validDraft()
    d.mode = 'smoke'
    d.environment = { ...d.environment, browser: 'webkit-playwright', channel: 'chromium', dirty: true, commit: 'abc', build: 'dev server' }
    d.fixtures = []
    d.raw = {}
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.desktopAdmitted).toBe(false)
    const problems = receipt.evidenceProblems.join('\n')
    for (const needle of ['mode smoke', 'browser webkit-playwright', 'channel chromium', 'dirty', 'not a full SHA', 'build', 'fixture F2 missing', 'scroll.ordinary: no raw timeline', 'ack.sort: no raw records', 'import.F1: no raw runs', 'memory.F2: no completed session']) expect(problems).toContain(needle)
    expect(receipt.evaluation.filter(r => r.profile === 'desktop' && !/sort-|filter-1m|index/.test(r.id)).every(r => r.outcome === 'not-run' || r.outcome === 'pending')).toBe(true)
    expect(await verifyReceipt(receipt)).toEqual({ ok: true }) // integrity alone is not evidence
    expect((await validateReceipt(receipt, ctx)).admitted).toBe(false)
  })

  it('(a2) a fixture whose content does not match the manifest blocks admission', async () => {
    const d = validDraft()
    d.fixtures = d.fixtures.map(f => (f.id === 'F2' ? { ...f, nnz: 1, contentSha256: 'f'.repeat(64) } : f))
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems.some(p => p.startsWith('fixture F2: contentSha256'))).toBe(true)
    expect(receipt.evidenceProblems.some(p => p.startsWith('fixture F2: nnz'))).toBe(true)
  })

  it('(b) one-sample series cannot pass-limit their way in; only the documented exceptions may be limited, and never below 10', async () => {
    const d = validDraft()
    const raw = d.raw as ReturnType<typeof validRaw> & Record<string, any>
    raw.scroll.ordinary.timeline = timeline(1, 3)
    raw.ack.sort.records = ackRecords(1, 17)
    raw.firstViewport.F1 = runs(1, 480)
    raw.firstViewport.F2 = runs(9, 600)
    raw.ack.filterApply.records = ackRecords(9, 33)
    raw.import.F3.runs = [importRun()]
    const receipt = await finalizeReceipt(d, ctx)
    const problems = receipt.evidenceProblems.join('\n')
    for (const needle of ['scroll.ordinary.inputToPaint: 1 samples, below the minimum 300', 'scroll.ordinary.frameInterval: 1 samples', 'ack.sort: 1 samples, below the minimum 30', 'firstViewport.F1: 1 samples, below the minimum 30', 'firstViewport.F2: 9 samples, below the minimum 10', 'ack.filter: 9 samples, below the minimum 10', 'import.F3: 1 complete runs, below 3']) expect(problems).toContain(needle)
    expect(receipt.desktopAdmitted).toBe(false)
  })

  it('(c) slow raw samples behind fast stored summaries: recomputation misses the row and validation rejects the receipt', async () => {
    const honest = await finalizeReceipt(validDraft(), ctx)
    // Tamper: raw ordinary presents now 1,000 ms after input, stored summaries and verdict kept fast, then re-hashed.
    const tampered = JSON.parse(JSON.stringify(honest)) as ViewerReceipt & { raw: any }
    tampered.raw.scroll.ordinary.timeline.presents = tampered.raw.scroll.ordinary.timeline.events.map((t: number) => t + 1000)
    const { hash: _hash, ...body } = tampered
    const { canonicalize, sha256Hex } = await import('../src/debug/harness/result')
    tampered.hash = await sha256Hex(canonicalize(body))
    expect(await verifyReceipt(tampered)).toEqual({ ok: true })
    const verdict = await validateReceipt(tampered, ctx)
    expect(verdict.admitted).toBe(false)
    expect(verdict.problems).toContain('stored measurements disagree with the raw records')
    expect(verdict.problems).toContain('stored evaluation disagrees with the raw records')
    // Finalizing the same raw records computes the miss itself.
    const recomputed = await finalizeReceipt({ ...validDraft(), raw: tampered.raw }, ctx)
    expect(recomputed.evaluation.find(r => r.id === 'warm-viewport-frame' && r.profile === 'desktop')!.outcome).toBe('miss')
  })

  it('(d) censored acknowledgments count as misses, not as missing samples', async () => {
    const d = validDraft()
    const raw = d.raw as any
    // Two of 30 censored: the nearest-rank p95 (29th value) is a censored one.
    for (const i of [3, 7]) raw.ack.sort.records[i] = { ...raw.ack.sort.records[i], presentAt: null, what: null }
    const receipt = await finalizeReceipt(d, ctx)
    const row = receipt.evaluation.find(r => r.id === 'input-acknowledgment' && r.profile === 'desktop')!
    expect(receipt.measurements.ack.sort).toMatchObject({ samples: 30, max: ACK_CENSORED_MS })
    expect(row.outcome).toBe('miss')
  })

  it('(e) browser loss or a stalled memory session is a run error', async () => {
    const d = validDraft()
    const raw = d.raw as any
    raw.errors = ['browser disconnected']
    raw.memory.F4 = { stalled: true }
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems).toContain('run error: browser disconnected')
    expect(receipt.evidenceProblems).toContain('memory.F4: no completed session')
  })
})

describe('receipt hash', () => {
  it('detects any modification of a finalized receipt', async () => {
    const receipt = await finalizeReceipt(validDraft(), ctx)
    expect(await verifyReceipt(receipt)).toEqual({ ok: true })
    expect((await verifyReceipt({ ...receipt, mode: 'smoke' })).ok).toBe(false)
  })
})
