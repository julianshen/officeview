/**
 * S10.2 real-viewer performance suite (driven by scripts/bench-interactive-viewer.ts
 * --suite viewer / --fixture all).
 *
 * Measures the shipped viewer in a production build of src/debug/viewer-perf.html
 * (vite.perf.config.ts) served cross-origin isolated from loopback, in one
 * automated browser. Every scenario opens true XLSX packages from local Files:
 *   first-viewport  F1/F2/F3, fresh page per sample (module cache warm, session cold)
 *   scroll          ordinary (F2), wrapped, styled: 300 open-loop frames each,
 *                   plus a wide (F3 horizontal) and a real-wheel diagnostic
 *   ack             sort / filter open / find / cancel on F2 while the worker runs
 *   import          the paged importer alone in a worker, F1-F4
 *   memory          F2 and F4 sessions: baseline, peak, quiescent, post-close
 * Statistics and threshold evaluation live in src/debug/harness/viewerPerf.ts.
 * Nothing here asserts on wall clock; a release run decides admission from the
 * receipt's evaluation rows.
 */
import { chromium, webkit, type Browser, type BrowserContext, type CDPSession, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { cpus, loadavg, totalmem, freemem } from 'node:os'
import { join, resolve } from 'node:path'
import {
  ACK_CENSORED_MS, RELEASE_IDENTITY, finalizeReceipt, importThroughput, pairInputsToPresents, snapshotCodeBytes, snapshotTypeBytes, summarize,
  type ImportEvent, type ReceiptDraft, type SeriesSummary, type ViewerReceipt,
} from '../../src/debug/harness/viewerPerf'
import { MIN_FRAME_SAMPLES, MIN_OPERATION_SAMPLES } from '../../src/debug/harness/result'
import type { AckKind, AckRecord, ScrollRunResult, ViewerPerf } from '../../src/debug/viewer-perf'
import type { PerfReply } from '../../src/debug/viewerPerfWorker'
import type { BenchArgs } from './args'
import type { HarnessIdentity } from './identity'
import { startSampler } from './sampler'
import { measureSortOutcome, recordSortOutcome, type SortOutcome } from './viewerSortOutcome'
import { startStaticServer } from './static-server'
import { measureCodeSpace } from './codeSpace'
import { FIXTURE_MANIFEST_PATH, ensureViewerFixtures, type FixtureManifest, type ViewerFixture, type ViewerFixtureKey } from './viewerFixtures'

const PAGE = '/src/debug/viewer-perf.html'
const SCROLL_FRAMES = MIN_FRAME_SAMPLES
const SCROLL_WARMUP = 60
const ACK_TIMEOUT_MS = ACK_CENSORED_MS
const OPEN_TIMEOUT_MS = 180_000
/** Expensive 1M cold opens take fewer samples, labelled (spec §12.6). */
const F2_VIEWPORT_SAMPLES = 10
/** Sort acknowledgment stops early once this many samples all exceed 10x the ceiling. */
const EARLY_MISS_SAMPLES = 5
/** Filter apply + clear re-admits the 1M-row range per sample: fewer, labelled samples. */
const FILTER_APPLY_SAMPLES = 10
const MEMORY_SESSION_LIMIT_MS = 10 * 60_000
const RECOVERY_CYCLES = 10
const WARM_HOST_SAMPLES = 10
const WARM_OPEN_SAMPLES = 5
const CHECKPOINT_READINGS = 5
const WARM_NOISE_ROUNDS = 3
const LIVE_SAMPLE_MS = 300

type W = Window & { __viewerPerf: ViewerPerf }
/** Runs a closure-free function against window.__viewerPerf in the page. */
const perf = <T>(page: Page, fn: (p: ViewerPerf) => T | Promise<T>): Promise<T> =>
  page.evaluate(`Promise.resolve((${fn.toString()})(window.__viewerPerf))`) as Promise<T>

export interface SuiteContext {
  args: BenchArgs
  repoRoot: string
  label: string
  identity: HarnessIdentity
  log(message: string): void
}

function sh(command: string, argv: string[]): string {
  try { return execFileSync(command, argv, { encoding: 'utf8' }).trim() } catch { return 'unavailable' }
}

/** Resident set of every process descended from this one (the launched browser), in bytes. Informational. */
function browserProcessRss(): number | null {
  try {
    const rows = execFileSync('ps', ['-A', '-o', 'pid=,ppid=,rss='], { encoding: 'utf8' }).trim().split('\n').map(line => line.trim().split(/\s+/).map(Number))
    const children = new Map<number, number[]>()
    for (const [pid, ppid] of rows) children.set(ppid, [...(children.get(ppid) ?? []), pid])
    const rss = new Map(rows.map(([pid, , kb]) => [pid, kb * 1024]))
    let total = 0
    const stack = [...(children.get(process.pid) ?? [])]
    while (stack.length) { const pid = stack.pop()!; total += rss.get(pid) ?? 0; stack.push(...(children.get(pid) ?? [])) }
    return total
  } catch { return null }
}

function machine(): { machine: string; os: string; extra: Record<string, unknown> } {
  const cpu = process.platform === 'darwin' ? sh('sysctl', ['-n', 'machdep.cpu.brand_string']) : cpus()[0]?.model ?? 'unknown'
  const os = process.platform === 'darwin' ? `macOS ${sh('sw_vers', ['-productVersion'])} (${sh('sw_vers', ['-buildVersion'])})` : `${process.platform} ${process.arch}`
  const power = process.platform === 'darwin' ? sh('pmset', ['-g', 'batt']).split('\n').slice(0, 2).join(' | ') : 'unavailable'
  const thermal = process.platform === 'darwin' ? sh('pmset', ['-g', 'therm']).replace(/\s+/g, ' ') : 'unavailable'
  return { machine: `${cpu}, ${cpus().length} logical cores, ${Math.round(totalmem() / 2 ** 30)} GiB`, os, extra: { power, thermal, freeMemBytesAtStart: freemem() } }
}

async function newPerfPage(context: BrowserContext, base: string, errors: string[]): Promise<Page> {
  const page = await context.newPage()
  page.on('pageerror', error => errors.push(`pageerror: ${error.message}`))
  // Every console error is a run error (the static server answers the favicon, the only former 404).
  page.on('console', message => {
    if (message.type() === 'error') errors.push(`console: ${message.text().slice(0, 300)}`) })
  await page.goto(`${base}${PAGE}`)
  await page.waitForFunction(() => (window as unknown as W).__viewerPerf?.booted === true, undefined, { timeout: 30_000 })
  return page
}

async function openFixture(page: Page, fixture: ViewerFixture): Promise<void> {
  await page.setInputFiles('#perf-file', fixture.path)
  await page.waitForFunction(() => { const p = (window as unknown as W).__viewerPerf; return p.marks.firstUsableAt !== null || p.error !== null }, undefined, { timeout: OPEN_TIMEOUT_MS })
  const error = await perf(page, p => p.error)
  if (error) throw new Error(`${fixture.key}: viewer error: ${error}`)
}

async function waitComplete(page: Page): Promise<boolean> {
  try {
    await page.waitForFunction(() => (window as unknown as W).__viewerPerf.marks.completeAt !== null, undefined, { timeout: OPEN_TIMEOUT_MS })
    return true
  } catch { return false }
}

const relMarks = (marks: ViewerPerf['marks']) => {
  const at = (t: number | null) => (t === null || marks.blobAt === null ? null : t - marks.blobAt)
  return { firstPresent: at(marks.firstPresentAt), firstUsable: at(marks.firstUsableAt), status: at(marks.statusAt), complete: at(marks.completeAt), usableStatus: marks.usableStatus }
}

// ---- scenarios ---------------------------------------------------------------

async function firstViewport(context: BrowserContext, base: string, fixture: ViewerFixture, samples: number, errors: string[], log: (m: string) => void) {
  // One discarded open warms the browser's module and code caches; every
  // measured sample is a fresh page (new worker, new session, new viewer).
  const warm = await newPerfPage(context, base, errors)
  await openFixture(warm, fixture)
  await warm.close()
  const runs: ReturnType<typeof relMarks>[] = []
  for (let i = 0; i < samples; i++) {
    const page = await newPerfPage(context, base, errors)
    await openFixture(page, fixture)
    await waitComplete(page)
    runs.push(relMarks(await perf(page, p => p.marks)))
    // Close the document first so its worker and page storage are released.
    await perf(page, p => p.close())
    await page.close()
  }
  const usable = runs.map(r => r.firstUsable!).filter(v => v !== null)
  log(`    first usable viewport ${fixture.key}: n=${usable.length} p50 ${fmt(summarize(usable, 1))}`)
  return { runs, usable }
}

async function scrollSession(context: BrowserContext, base: string, fixture: ViewerFixture, motion: { dx: number; dy: number }, errors: string[], wheel: boolean) {
  const page = await newPerfPage(context, base, errors)
  await openFixture(page, fixture)
  const complete = await waitComplete(page)
  await page.waitForTimeout(1000)
  const result: ScrollRunResult = await page.evaluate(({ warmup, frames, dx, dy }) => (window as unknown as W).__viewerPerf.scrollRun({ warmup, frames, dx, dy }), { warmup: SCROLL_WARMUP, frames: SCROLL_FRAMES, ...motion })
  let wheelResult: { latencies: number[]; coalesced: number; unmatched: number; events: number; presents: number } | null = null
  if (wheel) {
    // Real input through the browser: CDP wheel events over the grid, back to back.
    await page.mouse.move(720, 480)
    await perf(page, p => p.wheelRecord(true))
    for (let i = 0; i < SCROLL_FRAMES; i++) await page.mouse.wheel(0, i % 120 < 60 ? 60 : -60)
    await page.waitForTimeout(250)
    const log = await perf(page, p => p.wheelRecord(false))
    wheelResult = { ...pairInputsToPresents(log.wheel, log.presents), events: log.wheel.length, presents: log.presents.length }
  }
  const bitmap = await perf(page, p => ({ peak: p.peakBitmapBytes, budget: p.lastFrame?.budgetBytes ?? null, canvas: p.canvasBytes() }))
  await page.close()
  return { complete, result, wheel: wheelResult, bitmap }
}

async function ackSuite(context: BrowserContext, base: string, f2: ViewerFixture, repeat: number, errors: string[], log: (m: string) => void) {
  const page = await newPerfPage(context, base, errors)
  await openFixture(page, f2)
  await waitComplete(page)
  await page.waitForTimeout(500)
  // Arming completes (listeners installed) before the input is sent; the result is awaited after.
  const arm = async (kind: AckKind) => {
    await page.evaluate(({ kind, timeout }) => (window as unknown as W).__viewerPerf.ackArm(kind, timeout), { kind, timeout: ACK_TIMEOUT_MS })
    return () => page.evaluate(() => (window as unknown as W).__viewerPerf.ackWait())
  }
  const latency = (r: AckRecord) => (r.presentAt !== null && r.inputAt >= 0 ? r.presentAt - r.inputAt : null)
  type AckSeries = { samples: number[]; records: AckRecord[]; censored: number; note: string; endToEnd?: number[]; operations?: SortOutcome[] }

  // Sort: observe completion or refusal before clicking; acknowledgment and
  // successful operation latency are separate measurements.
  const sortButton = page.getByRole('button', { name: 'Cycle sort column G' })
  const sort: AckSeries & { operations: SortOutcome[]; endToEnd: number[] } = { samples: [], records: [], censored: 0, note: '', endToEnd: [], operations: [] }
  for (let i = 0; i < repeat; i++) {
    const wait = await arm('status')
    const operation = await measureSortOutcome(page, () => sortButton.click())
    const record = await wait()
    recordSortOutcome(sort, operation, errors)
    sort.records.push(record)
    const value = latency(record)
    if (value === null) sort.censored++
    else sort.samples.push(value)
    if (operation.outcome !== 'completed') {
      sort.note = `stopped after ${i + 1} samples: sort ${operation.outcome}; see operations and run errors`
      break
    }
    if (i + 1 === EARLY_MISS_SAMPLES && sort.records.every(r => { const v = latency(r); return v === null || v > 500 })) {
      sort.note = `stopped after ${EARLY_MISS_SAMPLES} samples: every acknowledgment exceeded 10x the 50 ms ceiling (no pending state before the result)`
      break
    }
  }
  log(`    ack sort: ${sort.samples.length} samples (${sort.censored} censored at ${ACK_TIMEOUT_MS} ms) p50 ${fmt(summarize(sort.samples, 1))}; completed end-to-end p50 ${fmt(summarize(sort.endToEnd, 1))}; outcomes ${sort.operations.map(r => r.outcome).join(', ')}`)

  // Filter open and Cancel while the value list loads in the worker.
  const funnel = page.getByRole('button', { name: /^Filter column D/ })
  const dialog = page.getByRole('dialog', { name: 'Column filter' })
  const filter: AckSeries = { samples: [] as number[], records: [] as AckRecord[], censored: 0, note: 'funnel click -> filter dialog presented' }
  const cancel: AckSeries = { samples: [] as number[], records: [] as AckRecord[], censored: 0, note: 'dialog Cancel while the column value list loads -> dialog removed' }
  for (let i = 0; i < repeat; i++) {
    const opened = await arm('dialog-open')
    await funnel.click()
    const r1 = await opened()
    filter.records.push(r1)
    const v1 = latency(r1)
    if (v1 === null) filter.censored++
    else filter.samples.push(v1)
    await dialog.waitFor({ state: 'visible', timeout: 30_000 })
    const closed = await arm('dialog-close')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    const r2 = await closed()
    cancel.records.push(r2)
    const v2 = latency(r2)
    if (v2 === null) cancel.censored++
    else cancel.samples.push(v2)
    await dialog.waitFor({ state: 'hidden', timeout: 30_000 })
  }
  log(`    ack filter-open: p50 ${fmt(summarize(filter.samples, 1))}; cancel p50 ${fmt(summarize(cancel.samples, 1))}`)

  // Filter apply: OK on a one-value selection; the dialog closes (or a status /
  // aria change) while the worker evaluates 1M rows. Each sample also clears.
  const apply: AckSeries = { samples: [] as number[], records: [] as AckRecord[], censored: 0, note: `OK on a value selection -> first qualifying change; ${FILTER_APPLY_SAMPLES} samples (each apply + clear re-admits the 1M-row range)`, endToEnd: [] as number[] }
  for (let i = 0; i < Math.min(repeat, FILTER_APPLY_SAMPLES); i++) {
    await funnel.click()
    await dialog.waitFor({ state: 'visible', timeout: 30_000 })
    await dialog.getByRole('button', { name: 'Deselect All', exact: true }).click({ timeout: 60_000 })
    await dialog.getByRole('checkbox', { name: 'alpha', exact: true }).check({ timeout: 60_000 })
    const wait = await arm('apply')
    const t0 = Date.now()
    await dialog.getByRole('button', { name: 'OK' }).click()
    const record = await wait()
    await page.waitForFunction(() => /Visible body \d+ \/ complete body/.test(document.querySelector('[data-officeview-ui=viewer-status]')?.textContent ?? '') && !/Visible body (\d+) \/ complete body \1\b/.test(document.querySelector('[data-officeview-ui=viewer-status]')?.textContent ?? ''), undefined, { timeout: 120_000 }).catch(() => {})
    apply.endToEnd!.push(Date.now() - t0)
    apply.records.push(record)
    const v = latency(record)
    if (v === null) apply.censored++
    else apply.samples.push(v)
    await funnel.click()
    await dialog.waitFor({ state: 'visible', timeout: 30_000 })
    await dialog.getByRole('button', { name: 'Clear all filters' }).click({ timeout: 60_000 })
    await page.waitForFunction(() => /Visible body (\d+) \/ complete body \1\b/.test(document.querySelector('[data-officeview-ui=viewer-status]')?.textContent ?? ''), undefined, { timeout: 120_000 }).catch(() => {})
    await dialog.waitFor({ state: 'hidden', timeout: 30_000 }).catch(() => {})
  }
  log(`    ack filter-apply: ${apply.samples.length} samples p50 ${fmt(summarize(apply.samples, 1))}; end-to-end p50 ${fmt(summarize(apply.endToEnd!, 1))}`)

  // Find: each keystroke while the previous query's search runs over the worker source.
  const box = page.getByLabel('Search document')
  const find: AckSeries = { samples: [] as number[], records: [] as AckRecord[], censored: 0, note: 'keystroke in Search document while the previous query searches -> echo frame' }
  await box.click()
  for (let i = 0; i < repeat; i++) {
    if (i % 10 === 0) await box.fill('')
    const wait = await arm('find')
    await page.keyboard.press(String.fromCharCode(97 + (i % 26)))
    const record = await wait()
    find.records.push(record)
    const v = latency(record)
    if (v === null) find.censored++
    else find.samples.push(v)
  }
  await box.fill('')
  log(`    ack find: p50 ${fmt(summarize(find.samples, 1))}`)
  const wire = await perf(page, p => p.wire.filter(w => w.kind !== 'viewport').slice(-2000))
  await page.close()
  return { sort, filterOpen: filter, filterApply: apply, cancel, find, wire }
}

async function importSuite(context: BrowserContext, base: string, fixtures: ViewerFixture[], profile: 'D' | 'M', runs: number, errors: string[], log: (m: string) => void) {
  const page = await newPerfPage(context, base, errors)
  const out: Record<string, { runs: { outcome: string; events: ImportEvent[]; throughput: ReturnType<typeof importThroughput> }[] }> = {}
  for (const fixture of fixtures) {
    await page.setInputFiles('#perf-import-file', fixture.path)
    const list: { outcome: string; events: ImportEvent[]; throughput: ReturnType<typeof importThroughput> }[] = []
    for (let i = 0; i < runs; i++) {
      const reply = await page.evaluate(profile => (window as unknown as W).__viewerPerf.importRun(profile), profile) as PerfReply
      if (reply.kind !== 'import') { list.push({ outcome: reply.kind === 'error' ? `error: ${reply.message}` : reply.kind, events: [], throughput: null }); continue }
      list.push({ outcome: reply.outcome, events: reply.events, throughput: reply.outcome === 'complete' ? importThroughput(reply.events) : null })
    }
    out[fixture.key] = { runs: list }
    const ok = list.map(r => r.throughput).filter((t): t is NonNullable<typeof t> => t !== null)
    log(`    import ${fixture.key}: ${ok.map(t => `${t.expandedMiBPerSec.toFixed(1)} MiB/s ${Math.round(t.cellsPerSec).toLocaleString('en-US')} cells/s`).join(' | ')}`)
  }
  await page.close()
  return out
}

interface Heap { usedSize: number; totalSize: number; embedderHeapUsedSize?: number; backingStorageSize?: number }
const heapBytes = (h: Heap) => h.usedSize + (h.embedderHeapUsedSize ?? 0) + (h.backingStorageSize ?? 0)
type Uasm = { bytes?: number; breakdown?: { bytes: number; types: string[]; attribution: { scope: string; url: string }[] }[]; unavailable?: string }
function uasmSplit(u: Uasm) {
  if (!u.breakdown) return { available: false as const, reason: u.unavailable ?? 'no breakdown' }
  let window = 0, worker = 0, canvas = 0, other = 0
  for (const entry of u.breakdown) {
    if (entry.types.includes('Canvas')) canvas += entry.bytes
    else if (entry.attribution.some(a => a.scope === 'DedicatedWorkerGlobalScope')) worker += entry.bytes
    else if (entry.attribution.some(a => a.scope === 'Window')) window += entry.bytes
    else other += entry.bytes
  }
  return { available: true as const, total: u.bytes ?? 0, window, worker, canvas, other }
}

async function memorySession(context: BrowserContext, base: string, fixture: ViewerFixture, warmFixture: ViewerFixture, ops: 'f2' | 'f4', errors: string[], log: (m: string) => void) {
  const page = await newPerfPage(context, base, errors)
  let cdp: CDPSession | null = null
  try { cdp = await context.newCDPSession(page) } catch { cdp = null }
  const gcMain = async () => { if (cdp) await cdp.send('HeapProfiler.collectGarbage') }
  const heap = async (): Promise<Heap | null> => (cdp ? (await cdp.send('Runtime.getHeapUsage')) as Heap : null)
  // Every probe is bounded: a probe that does not answer is recorded as a timeout
  // (named in the receipt), never a fabricated 0 and never a stalled run.
  const timeouts: string[] = []
  const bounded = async <T>(what: string, promise: Promise<T>, ms = 60_000): Promise<T | null> => {
    let timer: ReturnType<typeof setTimeout> | undefined
    const TIMEOUT = Symbol('timeout')
    const result = await Promise.race([promise, new Promise<typeof TIMEOUT>(resolve => { timer = setTimeout(() => resolve(TIMEOUT), ms) })]).finally(() => clearTimeout(timer))
    if (result === TIMEOUT) { timeouts.push(what); log(`      timeout: ${what}`); return null }
    return result
  }
  // Compiled code is engine memory the ledger does not own: each checkpoint
  // measures the heap snapshot's 'code' bytes (bytecode, machine code,
  // feedback), so the gate can exclude exactly the measured growth relative to
  // the warmed host (review round 4 decision). The snapshot runs after the
  // checkpoint's heap reading; with VIEWER_SUITE_LIVE_DIAG it is also kept.
  const snapshotStats = async (label: string): Promise<{ codeBytes: number | null; typeBytes: Record<string, number> } | null> => {
    if (!cdp) return null
    const chunks: string[] = []
    const take = (e: { chunk: string }) => chunks.push(e.chunk)
    cdp.on('HeapProfiler.addHeapSnapshotChunk', take)
    try { await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false }) } finally { cdp.off('HeapProfiler.addHeapSnapshotChunk', take) }
    const text = chunks.join('')
    if (process.env.VIEWER_SUITE_LIVE_DIAG) writeFileSync(`${process.env.VIEWER_SUITE_LIVE_DIAG}-${fixture.key}-${label}.heapsnapshot`, text)
    const snap = JSON.parse(text)
    return { codeBytes: snapshotCodeBytes(snap), typeBytes: snapshotTypeBytes(snap) }
  }
  // Compiled-code bytes for the gate: V8 per-space statistics from a
  // memory-infra dump (no heap snapshot, so no DOM node ids or number-string
  // cache entries are created by the measurement itself).
  const codeBytes = async (label: string): Promise<number | null> => {
    if (!cdp) return null
    const h = await heap()
    if (!h) return null
    const reading = await measureCodeSpace(cdp, h.usedSize)
    if (!reading) timeouts.push(`${label}: code-space dump did not match the page's heap`)
    return reading?.codeBytes ?? null
  }
  const diagSnapshot = async (label: string) => {
    if (!process.env.VIEWER_SUITE_LIVE_DIAG || !cdp) return
    // Diagnostic only (never in release): a heap snapshot at this point.
    const chunks: string[] = []
    const take = (e: { chunk: string }) => chunks.push(e.chunk)
    cdp.on('HeapProfiler.addHeapSnapshotChunk', take)
    await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false })
    cdp.off('HeapProfiler.addHeapSnapshotChunk', take)
    writeFileSync(`${process.env.VIEWER_SUITE_LIVE_DIAG}-${fixture.key}-${label}.heapsnapshot`, chunks.join(''))
  }
  // The checkpoint procedure's repeated readings: k times three full
  // collections 200 ms apart, then one heap reading. The same procedure runs on
  // the idle closed warm host to measure its noise independently (review r5).
  const repeatedReadings = async (label: string): Promise<number[]> => {
    const out: number[] = []
    for (let r = 0; r < CHECKPOINT_READINGS; r++) {
      for (let k = 0; k < 3; k++) { await bounded(`${label} gc ${r}.${k}`, gcMain()); await new Promise(resolve => setTimeout(resolve, 200)) }
      const h = await bounded(`${label} heap ${r}`, heap())
      if (h) out.push(heapBytes(h))
    }
    return out
  }
  // The background samplers (live forced-GC samples, no-GC/worker polls) pause
  // while a checkpoint measures: their own page calls allocate and would show
  // up as checkpoint reading spread (review r5: diagnosed as harness noise).
  let samplersPaused = false, samplersInFlight = 0
  const sampled = (probe: () => Promise<void>) => async () => {
    if (samplersPaused) return
    samplersInFlight++
    try { await probe() } finally { samplersInFlight-- }
  }
  const pauseSamplers = async () => {
    samplersPaused = true
    const started = Date.now()
    while (samplersInFlight > 0 && Date.now() - started < 30_000) await new Promise(resolve => setTimeout(resolve, 20))
  }
  const checkpoint = async (name: string) => {
    log(`      checkpoint ${fixture.key}/${name} at ${new Date().toISOString().slice(11, 19)}`)
    await pauseSamplers()
    try { return await measureCheckpoint(name) } finally { samplersPaused = false }
  }
  const measureCheckpoint = async (name: string) => {
    // At rest: nothing presented for 1 s and no worker request open.
    const quiet = await bounded(`${name} quiet`, perf(page, p => p.quiet(1000, 20_000)), 30_000)
    const instrumentation = await bounded(`${name} drop instrumentation`, perf(page, p => p.dropInstrumentation()))
    // Three full collections with pauses: lifetime charges end in
    // FinalizationRegistry callbacks, which run as tasks after a collection.
    await bounded(`${name} worker gc`, perf(page, p => p.workerGc()))
    for (let k = 0; k < 3; k++) { await bounded(`${name} main gc ${k}`, gcMain()); await new Promise(resolve => setTimeout(resolve, 200)) }
    const heapReadings = await repeatedReadings(name)
    const h = await bounded(`${name} heap`, heap())
    const sortedReadings = [...heapReadings].sort((x, y) => x - y)
    const medianHeap = sortedReadings.length ? sortedReadings[Math.floor(sortedReadings.length / 2)] : null
    const worker = await bounded(`${name} worker stats`, perf(page, p => p.workerStats()))
    const uasm = await bounded(`${name} uasm`, perf(page, p => p.measureMemory()))
    const canvas = (await bounded(`${name} canvas`, perf(page, p => p.canvasBytes()))) ?? { count: -1, bytes: -1 }
    const storage = (await bounded(`${name} storage`, perf(page, p => p.storageEstimate()), 15_000)) ?? { usage: null, quota: null }
    const frame = (await bounded(`${name} frame`, perf(page, p => ({ peakBitmapBytes: p.peakBitmapBytes, bitmapBytes: p.lastFrame?.bitmapBytes ?? null, budgetBytes: p.lastFrame?.budgetBytes ?? null, workersCreated: p.workersCreated, workersTerminated: p.workersTerminated }))))
      ?? { peakBitmapBytes: 0, bitmapBytes: null, budgetBytes: null, workersCreated: -1, workersTerminated: -1 }
    const ledger = await bounded(`${name} main ledger`, perf(page, p => p.mainLedger()))
    const code = await bounded(`${name} code bytes`, codeBytes(`cp-${name}`), 180_000)
    const layoutShifts = await bounded(`${name} layout shifts`, perf(page, p => p.layoutShifts()))
    const layoutShiftSources = await bounded(`${name} layout-shift sources`, perf(page, p => p.layoutShiftSources()))
    return { name, instrumentation, heap: h, heapBytes: medianHeap, heapReadings, codeBytes: code, layoutShifts, layoutShiftSources, quiet, ledger, worker, uasm: uasmSplit((uasm ?? { unavailable: 'probe timed out' }) as Uasm), canvas, storage, frame, browserProcessRssBytes: browserProcessRss() }
  }
  // Warmed host, established independently of the measured session: the same
  // page opens a small paged workbook (F2 at smoke scale), runs the same code
  // paths (scroll, header sort, filter dialog), closes it, and only then is
  // sampled. The samples' median is the attribution baseline; their spread is
  // the forced-GC probe noise, the only tolerance the cross-check allows.
  // The empty page before anything opens: how much the warm-up itself leaves
  // in the baseline (code, caches) is recorded, not assumed (review round 4).
  const warmEmptySamples: number[] = []
  for (let i = 0; i < WARM_OPEN_SAMPLES; i++) {
    for (let k = 0; k < 3; k++) { await bounded('empty gc', gcMain()); await new Promise(resolve => setTimeout(resolve, 100)) }
    const h = await bounded('empty heap', heap())
    if (h) warmEmptySamples.push(heapBytes(h))
  }
  // The empty page's compiled code and heap-snapshot totals by node type, so
  // the warm-up's own residue is measured by kind, not assumed (review r5).
  const emptySnapshot = await bounded('empty snapshot', snapshotStats('empty'), 180_000)
  await openFixture(page, warmFixture)
  await waitComplete(page)
  await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 30, dx: 0, dy: 60 }))
  {
    const header = page.getByRole('columnheader', { name: /^Column G/ })
    const before = await header.getAttribute('aria-sort')
    await page.getByRole('button', { name: 'Cycle sort column G' }).click()
    await page.waitForFunction(({ before }) => {
      const h = [...document.querySelectorAll('[role=columnheader]')].find(x => /^Column G/.test(x.getAttribute('aria-label') ?? x.textContent ?? ''))
      return (h?.getAttribute('aria-sort') ?? null) !== before
    }, { before }, { timeout: 120_000 }).catch(() => {})
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    await dialog.waitFor({ state: 'visible', timeout: 30_000 })
    await page.waitForTimeout(500)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
  }
  // The warm document still open after its operations, idle: forced-GC heap
  // next to the ledger's current bytes. Their gap above the closed warm host is
  // what an open viewer holds outside the ledger, measured before and apart
  // from the session, after the same code paths ran (so compiled code is in
  // both it and the baseline); it bounds the session's per-checkpoint gap
  // (review round 4).
  const warmOpen: { heapBytes: number; ledgerBytes: number }[] = []
  const warmOpenQuiet = await bounded('warm open quiet', perf(page, p => p.quiet(1000, 20_000)), 30_000)
  await bounded('warm open drop instrumentation', perf(page, p => p.dropInstrumentation()))
  for (let i = 0; i < WARM_OPEN_SAMPLES; i++) {
    for (let k = 0; k < 3; k++) { await bounded('warm open gc', gcMain()); await new Promise(resolve => setTimeout(resolve, 200)) }
    const h = await bounded('warm open heap', heap())
    const ledger = await bounded('warm open ledger', perf(page, p => p.mainLedger()?.bytes ?? null))
    if (h && typeof ledger === 'number') warmOpen.push({ heapBytes: heapBytes(h), ledgerBytes: ledger })
  }
  const warmOpenCode = await bounded('warm open code bytes', codeBytes('warm-open'), 180_000)
  const warmOpenLayoutShifts = await bounded('warm open layout shifts', perf(page, p => p.layoutShifts()))
  const warmOpenLayoutShiftSources = await bounded('warm open layout-shift sources', perf(page, p => p.layoutShiftSources()))
  await bounded('warm close', perf(page, p => p.close()), 30_000)
  // The type split for the warm-up residue check: a heap snapshot of the closed
  // warm host. Taken BEFORE the baseline readings, so whatever the snapshot
  // itself adds to page tables is in the baseline too; no snapshot is taken
  // during the measured session.
  const warmSnapshot = await bounded('warm snapshot', snapshotStats('warm'), 180_000)
  const warmHostSamples: number[] = []
  for (let i = 0; i < WARM_HOST_SAMPLES; i++) {
    for (let k = 0; k < 3; k++) { await bounded('warm gc', gcMain()); await new Promise(resolve => setTimeout(resolve, 100)) }
    const h = await bounded('warm heap', heap())
    if (h) warmHostSamples.push(heapBytes(h))
  }
  const warmCode = await bounded('warm code bytes', codeBytes('warm'), 180_000)
  const warmLayoutShifts = await bounded('warm layout shifts', perf(page, p => p.layoutShifts()))
  const warmLayoutShiftSources = await bounded('warm layout-shift sources', perf(page, p => p.layoutShiftSources()))
  // Independent noise of the checkpoint procedure itself: the idle closed warm
  // host, measured WARM_NOISE_ROUNDS times with exactly the checkpoint's k readings.
  const warmCheckpointNoise: number[][] = []
  for (let round = 0; round < WARM_NOISE_ROUNDS; round++) warmCheckpointNoise.push(await repeatedReadings(`warm noise ${round}`))
  const baseline = await checkpoint('baseline')
  // Peak sampling during open/import (no forced GC: an upper bound, garbage included).
  const polls: { t: number; heapBytes: number | null; workerPeak: number | null; workerRetained: number | null }[] = []
  // Forced-GC live samples through import, scroll, sort/value list and close:
  // the live heap at each instant. Their maximum is a LOWER bound of the live
  // peak (a spike between samples is missed); it can prove a miss, not a pass.
  const liveHeapSamples: number[] = []
  const failed = (what: string) => (error: unknown) => { timeouts.push(`${what} failed: ${error instanceof Error ? error.message : String(error)}`) }
  // Diagnostic only (VIEWER_SUITE_LIVE_DIAG, never in release): a second
  // collect + read right after each sample, and the ledger at that moment, to
  // tell a live spike from allocation between the collection and the read.
  const liveDiag: { t: number; first: number; second: number | null; ledger: number | null }[] = []
  const diagnose = Boolean(process.env.VIEWER_SUITE_LIVE_DIAG)
  const stopLive = startSampler(LIVE_SAMPLE_MS, sampled(async () => {
    await bounded('live gc', gcMain(), 10_000)
    const h = await bounded('live heap', heap(), 10_000)
    if (h) liveHeapSamples.push(heapBytes(h))
    if (h && diagnose) {
      await bounded('diag gc', gcMain(), 10_000)
      const h2 = await bounded('diag heap', heap(), 10_000)
      const ledger = await bounded('diag ledger', perf(page, p => p.mainLedger()?.bytes ?? null), 10_000)
      liveDiag.push({ t: Date.now(), first: heapBytes(h), second: h2 ? heapBytes(h2) : null, ledger })
    }
  }), failed('live sample'))
  // No-GC sampling (an upper bound, garbage included) and worker ledger polls.
  const stopPolls = startSampler(250, sampled(async () => {
    const [h, w] = await Promise.all([bounded('poll heap', heap(), 10_000), bounded('poll worker stats', perf(page, p => p.workerStats()), 10_000)])
    const session = w && w.kind === 'stats' ? w.sessions[0]?.memory : undefined
    polls.push({ t: Date.now(), heapBytes: h ? heapBytes(h) : null, workerPeak: session?.peakBytes ?? null, workerRetained: session?.retainedBytes ?? null })
  }), failed('poll'))
  let loaded: Awaited<ReturnType<typeof checkpoint>>, afterOps: Awaited<ReturnType<typeof checkpoint>>
  try {
  await openFixture(page, fixture)
  await waitComplete(page)
  loaded = await checkpoint('loaded')
  if (ops === 'f2') {
    await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 120, dx: 0, dy: 60 }))
    if (process.env.VIEWER_SUITE_LIVE_DIAG) { await gcMain(); await diagSnapshot('scroll') }
    const header = page.getByRole('columnheader', { name: /^Column G/ })
    const before = await header.getAttribute('aria-sort')
    await page.getByRole('button', { name: 'Cycle sort column G' }).click()
    await page.waitForFunction(({ before }) => {
      const h = [...document.querySelectorAll('[role=columnheader]')].find(x => /^Column G/.test(x.getAttribute('aria-label') ?? x.textContent ?? ''))
      return (h?.getAttribute('aria-sort') ?? null) !== before
    }, { before }, { timeout: 120_000 }).catch(() => {})
  } else {
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    await dialog.waitFor({ state: 'visible', timeout: 30_000 })
    await page.waitForTimeout(5000)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
  }
  afterOps = await checkpoint('after-ops')
  } finally {
    // Every exit path stops both samplers.
    stopLive()
    stopPolls()
  }
  await bounded('viewer close', perf(page, p => p.close()), 30_000)
  await new Promise(resolve => setTimeout(resolve, 500))
  const closed = await checkpoint('closed')
  if (process.env.VIEWER_SUITE_SNAPSHOT && cdp && ops === 'f2') {
    const chunks: string[] = []
    cdp.on('HeapProfiler.addHeapSnapshotChunk', (e: { chunk: string }) => chunks.push(e.chunk))
    await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false })
    writeFileSync(process.env.VIEWER_SUITE_SNAPSHOT, chunks.join(''))
  }
  // Repeated open/close (F2: open, scroll, header sort, close): post-GC main heap
  // owned after each close. A real retention grows per cycle; engine slack
  // (compiled code, script sources, timeline buffers) levels off.
  const cycles: (number | null)[] = []
  if (ops === 'f2') {
    for (let i = 0; i < RECOVERY_CYCLES; i++) {
      await openFixture(page, fixture)
      await waitComplete(page)
      await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 60, dx: 0, dy: 60 }))
      const header = page.getByRole('columnheader', { name: /^Column G/ })
      const before = await header.getAttribute('aria-sort')
      await page.getByRole('button', { name: 'Cycle sort column G' }).click()
      await page.waitForFunction(({ before }) => {
        const h = [...document.querySelectorAll('[role=columnheader]')].find(x => /^Column G/.test(x.getAttribute('aria-label') ?? x.textContent ?? ''))
        return (h?.getAttribute('aria-sort') ?? null) !== before
      }, { before }, { timeout: 120_000 }).catch(() => {})
      await bounded('cycle close', perf(page, p => p.close()), 30_000)
      for (let k = 0; k < 3; k++) { await bounded('cycle gc', gcMain()); await new Promise(resolve => setTimeout(resolve, 200)) }
      const h = await bounded('cycle heap', heap())
      cycles.push(h && baseline.heapBytes !== null ? heapBytes(h) - baseline.heapBytes : null)
    }
    log(`      recovery cycles (MiB owned after each close): ${cycles.map(v => (v === null ? 'n/a' : (v / 2 ** 20).toFixed(2))).join(' ')}`)
  }
  await page.close()
  const owned = (c: typeof baseline) => (c.heapBytes !== null && baseline.heapBytes !== null ? c.heapBytes - baseline.heapBytes : null)
  const ledgerPeak = Math.max(0, ...polls.map(p => p.workerPeak ?? 0), ...[loaded, afterOps].map(c => (c.worker && c.worker.kind === 'stats' ? c.worker.sessions[0]?.memory.peakBytes ?? 0 : 0)))
  const uasmWorker = Math.max(0, ...[loaded, afterOps].map(c => (c.uasm.available ? c.uasm.worker : 0)))
  const uasmCanvas = Math.max(0, ...[loaded, afterOps].map(c => (c.uasm.available ? c.uasm.canvas : 0)))
  const summary = {
    mainOwnedBytes: Math.max(...[loaded, afterOps].map(c => owned(c) ?? Number.NEGATIVE_INFINITY)),
    mainPeakNoGcBytes: baseline.heapBytes === null ? null : Math.max(0, ...polls.map(p => (p.heapBytes ?? 0) - baseline.heapBytes!)),
    workerLedgerPeakBytes: ledgerPeak,
    workerUasmBytes: loaded.uasm.available ? uasmWorker : null,
    workerOwnedBytes: Math.max(ledgerPeak, uasmWorker),
    bitmapLedgerPeakBytes: Math.max(loaded.frame.peakBitmapBytes, afterOps.frame.peakBitmapBytes),
    bitmapUasmCanvasBytes: loaded.uasm.available ? uasmCanvas : null,
    bitmapPeakBytes: Math.max(loaded.frame.peakBitmapBytes, afterOps.frame.peakBitmapBytes, uasmCanvas),
    browserProcessRss: { note: 'RSS of every launched browser process (all tabs/GPU/utility); informational, outside the app pools', baseline: baseline.browserProcessRssBytes, loaded: loaded.browserProcessRssBytes, afterOps: afterOps.browserProcessRssBytes, closed: closed.browserProcessRssBytes },
    recovery: {
      mainAfterCloseBytes: owned(closed),
      cyclesMainOwnedBytes: cycles,
      canvasAfterClose: closed.canvas,
      workersCreated: closed.frame.workersCreated,
      workersTerminated: closed.frame.workersTerminated,
      uasmTotalDeltaBytes: closed.uasm.available && baseline.uasm.available ? closed.uasm.total - baseline.uasm.total : null,
      storageDeltaBytes: closed.storage.usage !== null && baseline.storage.usage !== null ? closed.storage.usage - baseline.storage.usage : null,
    },
  }
  log(`    memory ${fixture.key}: main ledger peak ${mib(afterOps.ledger?.peakBytes ?? null)}; main owned ${mib(summary.mainOwnedBytes)} (forced-GC live max ${mib(liveHeapSamples.length && baseline.heapBytes !== null ? Math.max(...liveHeapSamples) - baseline.heapBytes : null)} over ${liveHeapSamples.length} samples; peak no-GC ${mib(summary.mainPeakNoGcBytes)}), worker ledger ${mib(ledgerPeak)} / UASM ${mib(summary.workerUasmBytes)}, bitmap ${mib(summary.bitmapPeakBytes)}; after close main ${mib(summary.recovery.mainAfterCloseBytes)}, workers ${summary.recovery.workersTerminated}/${summary.recovery.workersCreated} terminated, canvases ${closed.canvas.count}`)
  if (diagnose) writeFileSync(`${process.env.VIEWER_SUITE_LIVE_DIAG}-${fixture.key}.json`, JSON.stringify({ warm: warmHostSamples, liveDiag }))
  return { summary, warmHost: { fixture: warmFixture.key, scale: warmFixture.scale, samples: warmHostSamples, open: warmOpen, codeBytes: warmCode, openCodeBytes: warmOpenCode, layoutShifts: warmLayoutShifts, openLayoutShifts: warmOpenLayoutShifts, layoutShiftSources: warmLayoutShiftSources, openLayoutShiftSources: warmOpenLayoutShiftSources, checkpointNoise: warmCheckpointNoise, openQuiet: warmOpenQuiet, empty: warmEmptySamples, emptyCodeBytes: emptySnapshot?.codeBytes ?? null, emptyTypeBytes: emptySnapshot?.typeBytes ?? null, typeBytes: warmSnapshot?.typeBytes ?? null }, checkpoints: { baseline, loaded, afterOps, closed }, polls, liveHeapSamples, liveSampleMs: LIVE_SAMPLE_MS, timeouts }
}

// ---- formatting ----------------------------------------------------------------

const mib = (v: number | null) => (v === null || !Number.isFinite(v) ? 'n/a' : `${(v / 2 ** 20).toFixed(1)} MiB`)
function fmt(s: SeriesSummary): string {
  if (s.status === 'no-samples') return 'no samples'
  return `${s.p50.toFixed(1)} ms${s.status === 'ok' ? `, p95 ${s.p95.toFixed(1)} ms` : ` (n=${s.samples}, p95 refused)`}, max ${s.max.toFixed(1)} ms`
}

// ---- suite ---------------------------------------------------------------------

export async function runViewerSuite(ctx: SuiteContext): Promise<{ receipt: ViewerReceipt; path: string }> {
  const { args, repoRoot, log } = ctx
  const startedAt = new Date().toISOString()
  const load0 = loadavg()
  log('building the production perf page (vite.perf.config.ts)…')
  execFileSync('bunx', ['vite', 'build', '--config', 'vite.perf.config.ts'], { cwd: repoRoot, stdio: 'inherit' })
  log(`verifying ${args.scale} fixtures against ${FIXTURE_MANIFEST_PATH} (regenerating any mismatch)…`)
  const fixtureManifest = JSON.parse(readFileSync(join(repoRoot, FIXTURE_MANIFEST_PATH), 'utf8')) as FixtureManifest
  const fixtures = await ensureViewerFixtures(args.scale, fixtureManifest)
  // The memory sessions' warmed-host baseline opens the smoke-scale F2 first.
  const warmFixtures = await ensureViewerFixtures('reduced', fixtureManifest, ['F2'])
  for (const f of Object.values(fixtures)) log(`  ${f.key}: ${f.bytes.toLocaleString('en-US')} B, content ${f.contentSha256.slice(0, 16)}…, ${f.scale}`)
  const server = await startStaticServer(join(repoRoot, '.tmp', 'viewer-perf-dist'), args.port)
  const base = server.url
  const engine = args.browser === 'webkit' ? webkit : chromium
  const browser: Browser = await engine.launch({
    headless: true,
    ...(args.channel === 'chrome' && args.browser === 'chromium' ? { channel: 'chrome' } : {}),
    ...(args.browser === 'chromium' ? { args: ['--js-flags=--expose-gc'] } : {}),
  })
  const viewport = args.profile === 'desktop' ? { width: 1440, height: 900 } : { width: 1024, height: 768 }
  const contextOptions = { viewport, deviceScaleFactor: 2, isMobile: args.profile === 'mobile', hasTouch: args.profile === 'mobile' }
  const errors: string[] = []
  const raw: Record<string, unknown> = {}
  // An unexpected disconnect (browser crash or an external kill on a shared
  // machine) ends the run loudly instead of leaving awaits pending forever.
  let closing = false
  browser.on('disconnected', () => {
    if (closing) return
    log('  ERROR: the browser disconnected unexpectedly; no receipt is written')
    process.exit(3)
  })
  const fresh = async () => {
    const context = await browser.newContext(contextOptions)
    context.on('page', page => page.on('crash', () => { errors.push('page crashed'); log('  ERROR: a page crashed') }))
    return context
  }
  // Diagnostic iteration only (never in release): VIEWER_SUITE_ONLY=memory,ack …
  if (process.env.VIEWER_SUITE_HEARTBEAT) setInterval(() => log(`  heartbeat ${new Date().toISOString().slice(11, 19)}`), 30_000).unref()
  const only = args.mode === 'release' ? null : process.env.VIEWER_SUITE_ONLY?.split(',') ?? null
  if (args.mode === 'release') delete process.env.VIEWER_SUITE_LIVE_DIAG
  const want = (name: string) => only === null || only.includes(name)
  const profileId = args.profile === 'desktop' ? 'D' : 'M'
  try {
    if (want('firstViewport')) {
    log('first usable viewport (local File -> labelled usable frame)…')
    const fv: Record<string, unknown> = {}
    for (const key of ['F1', 'F3', 'F2'] as const) {
      const context = await fresh()
      const samples = key === 'F2' ? Math.min(args.repeat, F2_VIEWPORT_SAMPLES) : args.repeat
      const run = await firstViewport(context, base, fixtures[key], samples, errors, log)
      await context.close()
      fv[key] = { ...run, samplesRequested: samples, limitation: key === 'F2' ? `1M-row cold opens: ${samples} samples, p95 refused below ${MIN_OPERATION_SAMPLES} (spec §12.6); max reported against the ceiling` : null }
    }
    raw.firstViewport = fv
    }

    if (want('scroll')) {
    log('sustained scroll (300 open-loop frames per mix)…')
    const scroll: Record<string, unknown> = {}
    const mixes: [string, ViewerFixtureKey, { dx: number; dy: number }, boolean][] = [
      ['ordinary', 'F2', { dx: 0, dy: 60 }, true], ['wrapped', 'wrapped', { dx: 0, dy: 60 }, false], ['styled', 'styled', { dx: 0, dy: 60 }, false],
      // Informational (not threshold inputs): horizontal F3 and non-ASCII wrapped text.
      ['wide', 'F3', { dx: 120, dy: 0 }, false], ['wrapped-intl', 'wrappedIntl', { dx: 0, dy: 60 }, false],
    ]
    for (const [kind, key, motion, wheel] of mixes) {
      const context = await fresh()
      const session = await scrollSession(context, base, fixtures[key], motion, errors, wheel)
      await context.close()
      const inputToPaint = summarize(session.result.inputToPaint, MIN_FRAME_SAMPLES)
      const frameInterval = summarize(session.result.frameIntervals, MIN_FRAME_SAMPLES)
      scroll[kind] = { fixture: key, motion, importComplete: session.complete, ...session.result, inputToPaintSummary: inputToPaint, frameIntervalSummary: frameInterval, stepToPaintSummary: summarize(session.result.stepToPaint, MIN_FRAME_SAMPLES), wheel: session.wheel && { ...session.wheel, summary: summarize(session.wheel.latencies, MIN_FRAME_SAMPLES) }, bitmap: session.bitmap }
      log(`    scroll ${kind}: input-to-paint ${fmt(inputToPaint)}; frame interval ${fmt(frameInterval)}; ${session.result.coalesced} coalesced, ${session.result.longTasks.length} long tasks${session.wheel ? `; wheel ${fmt(summarize(session.wheel.latencies, 1))}` : ''}`)
    }
    raw.scroll = scroll
    }

    if (want('ack')) {
    log('input acknowledgment on F2 while the worker runs…')
    {
      const context = await fresh()
      const ack = await ackSuite(context, base, fixtures.F2, args.repeat, errors, log)
      await context.close()
      raw.ack = ack
    }
    }

    if (want('import')) {
    log('importer throughput (paged importer alone, in a worker)…')
    {
      const context = await fresh()
      const imp = await importSuite(context, base, [fixtures.F1, fixtures.F2, fixtures.F3, fixtures.F4], profileId, 3, errors, log)
      await context.close()
      raw.import = imp
    }
    }

    if (want('memory')) {
    log('memory (ledger + CDP heap + measureUserAgentSpecificMemory), F2 and F4 sessions…')
    {
      const memory: Record<string, unknown> = {}
      for (const [key, ops] of [['F2', 'f2'], ['F4', 'f4']] as const) {
        const context = await fresh()
        // Watchdog: an intermittent stall of this session was seen on the shared
        // machine; a stalled session is recorded (memory row not-run), never hung.
        let timer: ReturnType<typeof setTimeout> | undefined
        const stalled = new Promise<null>(resolve => { timer = setTimeout(() => resolve(null), MEMORY_SESSION_LIMIT_MS) })
        const session = await Promise.race([memorySession(context, base, fixtures[key], warmFixtures.F2, ops, errors, log).catch((error: unknown) => { errors.push(`memory ${key}: ${error instanceof Error ? error.message : String(error)}`); return null }), stalled]).finally(() => clearTimeout(timer))
        await context.close().catch(() => {})
        if (session === null) {
          errors.push(`memory ${key}: session exceeded ${MEMORY_SESSION_LIMIT_MS / 60_000} min and was abandoned`)
          log(`  ERROR: memory ${key} session exceeded ${MEMORY_SESSION_LIMIT_MS / 60_000} min; recorded not-run`)
          memory[key] = { stalled: true }
          continue
        }
        memory[key] = session
      }
      raw.memory = memory
    }
    }
  } finally {
    closing = true
    await browser.close().catch(() => {})
    await server.close()
  }
  raw.errors = errors.slice(0, 200)
  const env = machine()
  const draft: ReceiptDraft = {
    schema: 'officeview-viewer-perf/1',
    mode: args.mode,
    environment: {
      commit: ctx.identity.commit, dirty: ctx.identity.dirty, browser: args.browser === 'webkit' ? 'webkit-playwright' : args.channel === 'chrome' ? 'chrome' : 'chromium',
      browserVersion: browser.version(), channel: args.channel, headless: true, profile: args.profile, viewport, dpr: 2,
      machine: env.machine, os: env.os, loadAverage: [...load0, ...loadavg()],
      contention: 'Shared development machine: other agents may run tests concurrently; loadAverage holds 1/5/15-minute averages at start then at end. Re-run on a quiet machine before admission.',
      startedAt, finishedAt: new Date().toISOString(), build: RELEASE_IDENTITY.build, playwright: 'channel-launched headless', ...env.extra,
    },
    fixtures: Object.values(fixtures).map(f => ({ id: f.key, file: f.path.split('/').pop()!, bytes: f.bytes, sha256: f.sha256, contentSha256: f.contentSha256, uncompressedBytes: f.uncompressedBytes, nnz: f.nnz, rows: f.rows, cols: f.cols, scale: f.scale })),
    raw,
  }
  const profileJson = JSON.parse(readFileSync(join(repoRoot, 'tests/fixtures/xlsx-viewer/performance-profile.json'), 'utf8'))
  // Every threshold input is derived from `raw`; identity is checked against the full-scale manifest.
  const receipt = await finalizeReceipt(draft, { profile: profileJson, fixtures: fixtureManifest.scales.full })
  const out = resolve(args.output)
  mkdirSync(out, { recursive: true })
  const path = join(out, `viewer-perf-${args.browser}-${args.profile}-${args.mode}-${ctx.label.slice(0, 12)}-${startedAt.replace(/[:.]/g, '-')}.json`)
  writeFileSync(path, JSON.stringify(receipt, null, 2) + '\n')
  return { receipt, path }
}
