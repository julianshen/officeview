/**
 * S10.2 real-browser performance receipt helpers: present detection, input to
 * present pairing, sample summaries, importer throughput windows, threshold
 * evaluation and the hash-tagged receipt. Pure functions; no wall clock.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  ACK_CENSORED_MS, RELEASE_IDENTITY, SCROLL_CENSORED_MS, classifyPresent, finalizeReceipt, importThroughput, pairInputsToPresents, summarize, validateReceipt, verifyReceipt,
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
/** One import of the fixture's real workload (manifest cells and expanded bytes) over `ms`. */
const importRun = (id: string) => {
  const f = manifest.scales.full[id], done = Math.floor(f.uncompressedBytes * 0.9999)
  const ms = f.uncompressedBytes / (60 * MiB) * 1000 // ~60 MiB/s, as measured
  return { outcome: 'complete', events: [
    { t: 0, phase: 'inventory', expandedBytes: 1000, populatedCells: 0, compressedRead: 0 },
    { t: ms / 3, phase: 'sheet', expandedBytes: Math.floor(done / 3), populatedCells: Math.floor(f.nnz / 3), compressedRead: 0 },
    { t: ms, phase: 'done', expandedBytes: done, populatedCells: f.nnz, compressedRead: 0 },
  ] }
}
const ledgerAt = (peakBytes: number, bytes = peakBytes) => ({ bytes, peakBytes, byPool: {}, atPeak: {} })
const checkpoint = (heapBytes: number, workerPeak: number, ledgerPeak: number) => ({
  heapBytes, ledger: ledgerAt(ledgerPeak), worker: { kind: 'stats', sessions: [{ memory: { peakBytes: workerPeak } }] },
  uasm: { available: true, worker: workerPeak, canvas: 40 * MiB }, frame: { peakBitmapBytes: 100 * MiB, workersCreated: 1, workersTerminated: 0 },
  canvas: { count: 2, bytes: 1 }, storage: { usage: 200 * MiB, quota: 1e10 },
})
/** Baseline 5 MiB; live peak 30 MiB (25 owned); post-close floor 4 MiB; ledger peak 22 MiB: 25 - 4 - 22 = -1, attributed. */
const memorySession = (cycles: number[] | null) => ({
  checkpoints: {
    baseline: { heapBytes: 5 * MiB, storage: { usage: 0, quota: 1e10 } },
    loaded: checkpoint(20 * MiB, 200 * MiB, 18 * MiB), afterOps: checkpoint(25 * MiB, 220 * MiB, 22 * MiB),
    closed: { heapBytes: 9 * MiB, ledger: ledgerAt(22 * MiB, 0), uasm: { available: true, worker: 0, canvas: 0 }, frame: { workersCreated: 1, workersTerminated: 1, peakBitmapBytes: 100 * MiB }, canvas: { count: 0, bytes: 0 }, storage: { usage: 0, quota: 1e10 } },
  },
  summary: { recovery: { cyclesMainOwnedBytes: cycles } },
  polls: [{ workerPeak: 210 * MiB }], timeouts: [], liveHeapSamples: [12 * MiB, 30 * MiB, 18 * MiB],
})
function validRaw(): Record<string, unknown> {
  return {
    scroll: Object.fromEntries(['ordinary', 'wrapped', 'styled'].map(k => [k, { importComplete: true, timeline: timeline(300, 3) }])),
    ack: { sort: { records: ackRecords(30, 17) }, filterApply: { records: ackRecords(10, 33) }, find: { records: ackRecords(30, 14) }, cancel: { records: ackRecords(30, 16) } },
    firstViewport: { F1: runs(30, 480), F2: runs(10, 600), F3: runs(30, 670) },
    import: Object.fromEntries(['F1', 'F2', 'F3'].map(k => [k, { runs: [importRun(k), importRun(k), importRun(k)] }])),
    memory: { F2: memorySession(Array.from({ length: 10 }, (_, i) => (4 + i * 0.05) * MiB)), F4: memorySession(null) },
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
const desktop = (r: ViewerReceipt, id: string) => r.evaluation.find(row => row.id === id && row.profile === 'desktop')!

describe('evaluation from raw records', () => {
  it('admits fully valid evidence: every row derived from raw, the main row on an attributed ledger peak', async () => {
    const receipt = await finalizeReceipt(validDraft(), ctx)
    expect(receipt.evidenceProblems).toEqual([])
    expect(desktop(receipt, 'warm-viewport-frame')).toMatchObject({ outcome: 'pass' })
    expect(desktop(receipt, 'warm-viewport-frame').value).toBeCloseTo(16.67, 2)
    expect(desktop(receipt, 'input-acknowledgment').outcome).toBe('pass-limited') // filter apply: documented 10-sample exception
    expect(desktop(receipt, 'progressive-first-viewport').outcome).toBe('pass-limited') // F2 cold opens: documented exception
    expect(desktop(receipt, 'numeric-sort-1m').source).toMatch(/S8\.6/)
    expect(desktop(receipt, 'import-expanded-throughput').outcome).toBe('pass')
    expect(desktop(receipt, 'worker-owned-memory').value).toBe(220 * MiB)
    expect(desktop(receipt, 'main-owned-memory')).toMatchObject({ outcome: 'pass', value: 22 * MiB })
    expect(receipt.measurements.memory.mainUnattributedBytes).toBe(-1 * MiB)
    expect(receipt.desktopAdmitted).toBe(true)
    for (const row of receipt.evaluation.filter(r => r.profile === 'mobile')) expect(row.outcome).toBe('not-run')
    expect(await validateReceipt(receipt, ctx)).toEqual({ admitted: true, problems: [] })
  })

  it('a ledger that misses live heap beyond the tolerance leaves the row pending: unattributed main memory', async () => {
    const d = validDraft()
    ;(d.raw.memory as any).F4.liveHeapSamples = [40 * MiB] // 35 owned - floor 4 - ledger 22 = 9 MiB unattributed
    const receipt = await finalizeReceipt(d, ctx)
    expect(desktop(receipt, 'main-owned-memory').outcome).toBe('pending')
    expect(desktop(receipt, 'main-owned-memory').note).toMatch(/unattributed main memory/)
    expect(receipt.evidenceProblems.some(p => /unattributed main memory \(9 MiB/.test(p))).toBe(true)
    expect(receipt.desktopAdmitted).toBe(false)
  })

  it('within the tolerance the ledger peak is the admitted value', async () => {
    const d = validDraft()
    ;(d.raw.memory as any).F4.liveHeapSamples = [32.5 * MiB] // 27.5 - 4 - 22 = 1.5 MiB <= 2 MiB
    const receipt = await finalizeReceipt(d, ctx)
    expect(desktop(receipt, 'main-owned-memory').outcome).toBe('pass')
    expect(receipt.desktopAdmitted).toBe(true)
  })

  it('a ledger peak or a sampled live heap above 64 MiB is a miss', async () => {
    const over = validDraft()
    ;(over.raw.memory as any).F2.checkpoints.afterOps.ledger = ledgerAt(70 * MiB)
    expect(desktop(await finalizeReceipt(over, ctx), 'main-owned-memory').outcome).toBe('miss')
    const live = validDraft()
    ;(live.raw.memory as any).F2.liveHeapSamples = [100 * MiB]
    expect(desktop(await finalizeReceipt(live, ctx), 'main-owned-memory').outcome).toBe('miss')
  })
})

describe('adversarial receipts (review rounds 1 and 2)', () => {
  it('(a) smoke mode, dirty WebKit identity, bad commit, no fixtures and no raw data never admit', async () => {
    const d = validDraft()
    d.mode = 'smoke'
    d.environment = { ...d.environment, browser: 'webkit-playwright', channel: 'chromium', dirty: true, commit: 'abc', build: 'dev server' }
    d.fixtures = []
    d.raw = {}
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.desktopAdmitted).toBe(false)
    const problems = receipt.evidenceProblems.join('\n')
    for (const needle of ['mode smoke', 'browser webkit-playwright', 'channel chromium', 'dirty', 'not a full SHA', 'build', 'fixture F2 missing', 'scroll.ordinary: no raw timeline', 'ack.sort: no raw records', 'import.F1: no raw runs', 'memory.F2: no completed session', 'run error record missing']) expect(problems).toContain(needle)
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
    const raw = d.raw as any
    raw.scroll.ordinary.timeline = timeline(1, 3)
    raw.ack.sort.records = ackRecords(1, 17)
    raw.firstViewport.F1 = runs(1, 480)
    raw.firstViewport.F2 = runs(9, 600)
    raw.ack.filterApply.records = ackRecords(9, 33)
    raw.import.F3.runs = [importRun('F3')]
    const problems = (await finalizeReceipt(d, ctx)).evidenceProblems.join('\n')
    for (const needle of ['scroll.ordinary.inputToPaint: 1 samples, below the minimum 300', 'scroll.ordinary.frameInterval: 1 samples', 'ack.sort: 1 samples, below the minimum 30', 'firstViewport.F1: 1 samples, below the minimum 30', 'firstViewport.F2: 9 samples, below the minimum 10', 'ack.filter: 9 samples, below the minimum 10', 'import.F3: 1 valid complete runs, below 3']) expect(problems).toContain(needle)
  })

  it('(c) slow raw samples behind fast stored summaries: recomputation misses the row and validation rejects the receipt', async () => {
    const honest = await finalizeReceipt(validDraft(), ctx)
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
    const recomputed = await finalizeReceipt({ ...validDraft(), raw: tampered.raw }, ctx)
    expect(desktop(recomputed, 'warm-viewport-frame').outcome).toBe('miss')
  })

  it('(d) censored acknowledgments count as misses, not as missing samples', async () => {
    const d = validDraft()
    const raw = d.raw as any
    // Two of 30 censored: the nearest-rank p95 (29th value) is a censored one.
    for (const i of [3, 7]) raw.ack.sort.records[i] = { ...raw.ack.sort.records[i], presentAt: null, what: null }
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.measurements.ack.sort).toMatchObject({ samples: 30, max: ACK_CENSORED_MS })
    expect(desktop(receipt, 'input-acknowledgment').outcome).toBe('miss')
  })

  it('(r2-1a) reversed frame timestamps are rejected, never a negative p95', async () => {
    const d = validDraft()
    const t = (d.raw as any).scroll.wrapped.timeline
    t.frames = [...t.frames].reverse()
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems).toContain('scroll.wrapped: frame timestamps are not strictly increasing')
    expect(receipt.measurements.scroll.wrapped).toBeUndefined()
    expect(desktop(receipt, 'warm-viewport-frame').outcome).toBe('not-run')
  })

  it('(r2-1b) an acknowledgment presented before its input is censored and named', async () => {
    const d = validDraft()
    const rec = (d.raw as any).ack.find.records
    for (const i of [0, 1]) rec[i] = { ...rec[i], presentAt: rec[i].inputAt - 5, mutationAt: rec[i].inputAt - 6 }
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems).toContain('ack.find[0]: input -> mutation -> present out of order')
    expect(receipt.measurements.ack.find).toMatchObject({ min: 14, max: ACK_CENSORED_MS })
    expect(desktop(receipt, 'input-acknowledgment').outcome).toBe('miss')
  })

  it('(r2-1c) scroll inputs that never reached a present are censored misses, not dropped', async () => {
    const d = validDraft()
    const t = (d.raw as any).scroll.ordinary.timeline
    const last = t.presents.at(-1)
    t.events = [...t.events, ...Array.from({ length: 100 }, (_, i) => last + 1 + i)]
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.measurements.scroll.ordinary!.inputToPaint).toMatchObject({ samples: 400, max: SCROLL_CENSORED_MS })
    expect(desktop(receipt, 'warm-viewport-frame').outcome).toBe('miss')
  })

  it('(r2-2) an import that does not hold the fixture workload is rejected (1 cell / 1 KiB / 0.001 ms)', async () => {
    const d = validDraft()
    const fake = { outcome: 'complete', events: [
      { t: 0, phase: 'inventory', expandedBytes: 0, populatedCells: 0, compressedRead: 0 },
      { t: 0.001, phase: 'done', expandedBytes: 1024, populatedCells: 1, compressedRead: 0 },
    ] }
    ;(d.raw as any).import.F2.runs = [fake, fake, fake]
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems).toContain(`import.F2[0]: 1 populated cells, the fixture holds ${manifest.scales.full.F2.nnz}`)
    expect(receipt.evidenceProblems).toContain('import.F2: 0 valid complete runs, below 3')
    expect(receipt.measurements.import.F2).toBeUndefined()
    const shrunk = validDraft()
    const ev = (shrunk.raw as any).import.F1.runs[0].events
    ev[0].expandedBytes = 0; ev[1].expandedBytes = 5; ev[2].expandedBytes = 10
    expect((await finalizeReceipt(shrunk, ctx)).evidenceProblems.some(p => /import\.F1\[0\]: 10 expanded bytes/.test(p))).toBe(true)
    const backwards = validDraft()
    ;(backwards.raw as any).import.F3.runs[1].events[1].populatedCells = 10 ** 9
    expect((await finalizeReceipt(backwards, ctx)).evidenceProblems).toContain('import.F3[1]: progress is not finite, nonnegative and monotonic')
  })

  it('(r2-3a) missing F4 worker/bitmap probes block admission; F2 maxima cannot stand in', async () => {
    const d = validDraft()
    for (const name of ['loaded', 'afterOps']) {
      const cp = (d.raw as any).memory.F4.checkpoints[name]
      delete cp.worker; delete cp.uasm; delete cp.frame
    }
    const receipt = await finalizeReceipt(d, ctx)
    const problems = receipt.evidenceProblems.join('\n')
    for (const needle of ['memory.F4.loaded: no worker ledger', 'memory.F4.afterOps: no measureUserAgentSpecificMemory reading', 'memory.F4.loaded: no bitmap ledger']) expect(problems).toContain(needle)
    expect(receipt.measurements.memory.workerOwnedBytes).toBeNull()
    expect(desktop(receipt, 'worker-owned-memory').outcome).toBe('not-run')
    expect(receipt.desktopAdmitted).toBe(false)
  })

  it('(r2-3b) a missing or unclean close/recovery leg blocks admission', async () => {
    const gone = validDraft()
    delete (gone.raw as any).memory.F2.checkpoints.closed
    ;(gone.raw as any).memory.F2.summary.recovery.cyclesMainOwnedBytes = null
    const p1 = (await finalizeReceipt(gone, ctx)).evidenceProblems.join('\n')
    expect(p1).toContain('memory.F2.closed: missing checkpoint')
    expect(p1).toContain('memory.F2: 0 recovery cycles, below 10')
    const dirty = validDraft()
    const closed = (dirty.raw as any).memory.F4.checkpoints.closed
    closed.frame.workersTerminated = 0; closed.canvas.count = 1; closed.ledger.bytes = 4096; closed.storage.usage = 10
    const p2 = (await finalizeReceipt(dirty, ctx)).evidenceProblems.join('\n')
    for (const needle of ['memory.F4.closed: workers 0/1 terminated', 'memory.F4.closed: 1 canvases remain', 'memory.F4.closed: main ledger holds 4096 B after close', 'memory.F4.closed: page storage not back to baseline']) expect(p2).toContain(needle)
    const leak = validDraft()
    ;(leak.raw as any).memory.F2.summary.recovery.cyclesMainOwnedBytes = Array.from({ length: 10 }, (_, i) => (4 + i) * MiB)
    expect((await finalizeReceipt(leak, ctx)).evidenceProblems.some(p => /main heap grows 7 MiB over 10 open\/close cycles/.test(p))).toBe(true)
  })

  it('(r2-3c) any recorded run error blocks admission, page errors included', async () => {
    const d = validDraft()
    ;(d.raw as any).errors = ['pageerror: TypeError: Cannot read properties of undefined']
    const receipt = await finalizeReceipt(d, ctx)
    expect(receipt.evidenceProblems).toContain('run error: pageerror: TypeError: Cannot read properties of undefined')
    expect(receipt.desktopAdmitted).toBe(false)
  })

  it('(e) a stalled memory session is missing evidence', async () => {
    const d = validDraft()
    ;(d.raw as any).memory.F4 = { stalled: true }
    expect((await finalizeReceipt(d, ctx)).evidenceProblems).toContain('memory.F4: no completed session')
  })
})

describe('receipt hash', () => {
  it('detects any modification of a finalized receipt', async () => {
    const receipt = await finalizeReceipt(validDraft(), ctx)
    expect(await verifyReceipt(receipt)).toEqual({ ok: true })
    expect((await verifyReceipt({ ...receipt, mode: 'smoke' })).ok).toBe(false)
  })
})

describe('retained S10.2 Chrome receipt', () => {
  it('the receipt the profile cites validates from its raw records to exactly the recorded verdict', async () => {
    const evidence = profile.viewerEvidence as { results: string; hash: string; commit: string; desktopAdmitted: boolean; evidenceProblems: string[]; rows: Record<string, { outcome: string; value: number | null }> }
    const receipt = JSON.parse(readFileSync(join(process.cwd(), evidence.results), 'utf8')) as ViewerReceipt
    expect(receipt.hash).toBe(evidence.hash)
    expect(receipt.environment.commit).toBe(evidence.commit)
    const verdict = await validateReceipt(receipt, ctx)
    expect(verdict.admitted).toBe(evidence.desktopAdmitted)
    if (evidence.desktopAdmitted) {
      expect(verdict.problems).toEqual([])
      expect(receipt.evidenceProblems).toEqual([])
      for (const [id, row] of Object.entries(evidence.rows)) {
        expect(desktop(receipt, id).outcome, id).toBe(row.outcome)
        expect(desktop(receipt, id).value, id).toBe(row.value)
        const threshold = profile.thresholds.find((t: { id: string }) => t.id === id)
        if (threshold) {
          expect(threshold.desktop.frozenMs, id).toBe(threshold.desktop.specMs)
          expect(threshold.desktop.outcome, id).toBe(row.outcome === 'miss' ? 'miss' : 'pass')
        }
      }
    } else {
      expect(verdict.problems.length).toBeGreaterThan(0)
    }
    for (const row of receipt.evaluation.filter(r => r.profile === 'mobile')) expect(row.outcome).toBe('not-run')
  })
})
