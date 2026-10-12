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
  finalizeReceipt, importThroughput, pairInputsToPresents, summarize,
  type ImportEvent, type ReceiptDraft, type ReceiptMeasurements, type SeriesSummary, type ViewerReceipt,
} from '../../src/debug/harness/viewerPerf'
import { MIN_FRAME_SAMPLES, MIN_OPERATION_SAMPLES } from '../../src/debug/harness/result'
import type { AckKind, AckRecord, ScrollRunResult, ViewerPerf } from '../../src/debug/viewer-perf'
import type { PerfReply } from '../../src/debug/viewerPerfWorker'
import type { BenchArgs } from './args'
import type { HarnessIdentity } from './identity'
import { startStaticServer } from './static-server'
import { ensureViewerFixtures, type ViewerFixture, type ViewerFixtureKey } from './viewerFixtures'

const PAGE = '/src/debug/viewer-perf.html'
const SCROLL_FRAMES = MIN_FRAME_SAMPLES
const SCROLL_WARMUP = 60
const ACK_TIMEOUT_MS = 15_000
const OPEN_TIMEOUT_MS = 180_000
/** Expensive 1M cold opens take fewer samples, labelled (spec §12.6). */
const F2_VIEWPORT_SAMPLES = 10
/** Sort acknowledgment stops early once this many samples all exceed 10x the ceiling. */
const EARLY_MISS_SAMPLES = 5
/** Filter apply + clear re-admits the 1M-row range per sample: fewer, labelled samples. */
const FILTER_APPLY_SAMPLES = 10

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
  page.on('console', message => { if (message.type() === 'error') errors.push(`console: ${message.text().slice(0, 300)}`) })
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
  type AckSeries = { samples: number[]; records: AckRecord[]; censored: number; note: string; endToEnd?: number[] }

  // Sort: header sort button, F2 column G (1M body rows). The result's aria-sort
  // change ends one sample; the acknowledgment is the first qualifying status /
  // aria change after the input.
  const sortHeader = page.getByRole('columnheader', { name: /^Column G/ })
  const sortButton = page.getByRole('button', { name: 'Cycle sort column G' })
  const sort: AckSeries = { samples: [] as number[], records: [] as AckRecord[], censored: 0, note: '', endToEnd: [] as number[] }
  for (let i = 0; i < repeat; i++) {
    const before = await sortHeader.getAttribute('aria-sort')
    const wait = await arm('status')
    const t0 = Date.now()
    await sortButton.click()
    const record = await wait()
    await page.waitForFunction(({ before }) => {
      const header = [...document.querySelectorAll('[role=columnheader]')].find(h => /^Column G/.test(h.getAttribute('aria-label') ?? h.textContent ?? ''))
      return (header?.getAttribute('aria-sort') ?? null) !== before
    }, { before }, { timeout: 120_000 }).catch(() => {})
    sort.endToEnd!.push(Date.now() - t0)
    sort.records.push(record)
    const value = latency(record)
    if (value === null) sort.censored++
    else sort.samples.push(value)
    if (i + 1 === EARLY_MISS_SAMPLES && sort.records.every(r => { const v = latency(r); return v === null || v > 500 })) {
      sort.note = `stopped after ${EARLY_MISS_SAMPLES} samples: every acknowledgment exceeded 10x the 50 ms ceiling (no pending state before the result)`
      break
    }
  }
  log(`    ack sort: ${sort.samples.length} samples (${sort.censored} censored at ${ACK_TIMEOUT_MS} ms) p50 ${fmt(summarize(sort.samples, 1))}; end-to-end p50 ${fmt(summarize(sort.endToEnd!, 1))}`)

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

async function memorySession(context: BrowserContext, base: string, fixture: ViewerFixture, ops: 'f2' | 'f4', errors: string[], log: (m: string) => void) {
  const page = await newPerfPage(context, base, errors)
  let cdp: CDPSession | null = null
  try { cdp = await context.newCDPSession(page) } catch { cdp = null }
  const gcMain = async () => { if (cdp) await cdp.send('HeapProfiler.collectGarbage') }
  const heap = async (): Promise<Heap | null> => (cdp ? (await cdp.send('Runtime.getHeapUsage')) as Heap : null)
  const checkpoint = async (name: string) => {
    await gcMain()
    await perf(page, p => p.workerGc())
    await page.waitForTimeout(200)
    await gcMain()
    const [h, worker, uasm, canvas, storage, frame] = await Promise.all([
      heap(), perf(page, p => p.workerStats()), perf(page, p => p.measureMemory()), perf(page, p => p.canvasBytes()), perf(page, p => p.storageEstimate()),
      perf(page, p => ({ peakBitmapBytes: p.peakBitmapBytes, bitmapBytes: p.lastFrame?.bitmapBytes ?? null, budgetBytes: p.lastFrame?.budgetBytes ?? null, workersCreated: p.workersCreated, workersTerminated: p.workersTerminated })),
    ])
    return { name, heap: h, heapBytes: h ? heapBytes(h) : null, worker, uasm: uasmSplit(uasm as Uasm), canvas, storage, frame, browserProcessRssBytes: browserProcessRss() }
  }
  const baseline = await checkpoint('baseline')
  // Peak sampling during open/import (no forced GC: an upper bound, garbage included).
  const polls: { t: number; heapBytes: number | null; workerPeak: number | null; workerRetained: number | null }[] = []
  let polling = true
  const poller = (async () => {
    while (polling) {
      const [h, w] = await Promise.all([heap(), perf(page, p => p.workerStats())])
      const session = w && w.kind === 'stats' ? w.sessions[0]?.memory : undefined
      polls.push({ t: Date.now(), heapBytes: h ? heapBytes(h) : null, workerPeak: session?.peakBytes ?? null, workerRetained: session?.retainedBytes ?? null })
      await new Promise(r => setTimeout(r, 250))
    }
  })()
  await openFixture(page, fixture)
  await waitComplete(page)
  const loaded = await checkpoint('loaded')
  if (ops === 'f2') {
    await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 120, dx: 0, dy: 60 }))
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
  const afterOps = await checkpoint('after-ops')
  polling = false
  await poller
  await perf(page, p => p.close())
  await page.waitForTimeout(500)
  const closed = await checkpoint('closed')
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
      canvasAfterClose: closed.canvas,
      workersCreated: closed.frame.workersCreated,
      workersTerminated: closed.frame.workersTerminated,
      uasmTotalDeltaBytes: closed.uasm.available && baseline.uasm.available ? closed.uasm.total - baseline.uasm.total : null,
      storageDeltaBytes: closed.storage.usage !== null && baseline.storage.usage !== null ? closed.storage.usage - baseline.storage.usage : null,
    },
  }
  log(`    memory ${fixture.key}: main owned ${mib(summary.mainOwnedBytes)} (peak no-GC ${mib(summary.mainPeakNoGcBytes)}), worker ledger ${mib(ledgerPeak)} / UASM ${mib(summary.workerUasmBytes)}, bitmap ${mib(summary.bitmapPeakBytes)}; after close main ${mib(summary.recovery.mainAfterCloseBytes)}, workers ${summary.recovery.workersTerminated}/${summary.recovery.workersCreated} terminated, canvases ${closed.canvas.count}`)
  return { summary, checkpoints: { baseline, loaded, afterOps, closed }, polls }
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
  log(`generating/hashing ${args.scale} fixtures…`)
  const fixtures = await ensureViewerFixtures(args.scale)
  for (const f of Object.values(fixtures)) log(`  ${f.key}: ${f.bytes.toLocaleString('en-US')} B, sha256 ${f.sha256.slice(0, 16)}…, ${f.scale}`)
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
  const measurements: ReceiptMeasurements = {}
  const fresh = () => browser.newContext(contextOptions)
  const profileId = args.profile === 'desktop' ? 'D' : 'M'
  try {
    log('first usable viewport (local File -> labelled usable frame)…')
    const fv: Record<string, unknown> = {}
    measurements.firstViewport = {}
    for (const key of ['F1', 'F3', 'F2'] as const) {
      const context = await fresh()
      const samples = key === 'F2' ? Math.min(args.repeat, F2_VIEWPORT_SAMPLES) : args.repeat
      const run = await firstViewport(context, base, fixtures[key], samples, errors, log)
      await context.close()
      fv[key] = { ...run, samplesRequested: samples, limitation: key === 'F2' ? `1M-row cold opens: ${samples} samples, p95 refused below ${MIN_OPERATION_SAMPLES} (spec §12.6); max reported against the ceiling` : null }
      measurements.firstViewport[key] = summarize(run.usable, MIN_OPERATION_SAMPLES)
    }
    raw.firstViewport = fv

    log('sustained scroll (300 open-loop frames per mix)…')
    const scroll: Record<string, unknown> = {}
    measurements.scroll = {}
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
      if (kind === 'ordinary' || kind === 'wrapped' || kind === 'styled') measurements.scroll[kind] = { inputToPaint, frameInterval }
      log(`    scroll ${kind}: input-to-paint ${fmt(inputToPaint)}; frame interval ${fmt(frameInterval)}; ${session.result.coalesced} coalesced, ${session.result.longTasks.length} long tasks${session.wheel ? `; wheel ${fmt(summarize(session.wheel.latencies, 1))}` : ''}`)
    }
    raw.scroll = scroll

    log('input acknowledgment on F2 while the worker runs…')
    {
      const context = await fresh()
      const ack = await ackSuite(context, base, fixtures.F2, args.repeat, errors, log)
      await context.close()
      raw.ack = ack
      measurements.ack = {
        sort: summarize(ack.sort.samples.concat(Array(ack.sort.censored).fill(ACK_TIMEOUT_MS)), MIN_OPERATION_SAMPLES),
        filter: summarize(ack.filterApply.samples.concat(Array(ack.filterApply.censored).fill(ACK_TIMEOUT_MS)), MIN_OPERATION_SAMPLES),
        find: summarize(ack.find.samples.concat(Array(ack.find.censored).fill(ACK_TIMEOUT_MS)), MIN_OPERATION_SAMPLES),
        cancel: summarize(ack.cancel.samples.concat(Array(ack.cancel.censored).fill(ACK_TIMEOUT_MS)), MIN_OPERATION_SAMPLES),
      }
    }

    log('importer throughput (paged importer alone, in a worker)…')
    {
      const context = await fresh()
      const imp = await importSuite(context, base, [fixtures.F1, fixtures.F2, fixtures.F3, fixtures.F4], profileId, 3, errors, log)
      await context.close()
      raw.import = imp
      measurements.import = {}
      for (const key of ['F1', 'F2', 'F3'] as const) {
        const ok = imp[key].runs.map(r => r.throughput).filter((t): t is NonNullable<typeof t> => t !== null)
        // Lower-bound targets: the worst of the runs is the reported value.
        if (ok.length === imp[key].runs.length && ok.length > 0) measurements.import[key] = { expandedMiBPerSec: Math.min(...ok.map(t => t.expandedMiBPerSec)), cellsPerSec: Math.min(...ok.map(t => t.cellsPerSec)) }
      }
    }

    log('memory (ledger + CDP heap + measureUserAgentSpecificMemory), F2 and F4 sessions…')
    {
      const memory: Record<string, unknown> = {}
      let main = Number.NEGATIVE_INFINITY, worker = 0, bitmap = 0
      for (const [key, ops] of [['F2', 'f2'], ['F4', 'f4']] as const) {
        const context = await fresh()
        const session = await memorySession(context, base, fixtures[key], ops, errors, log)
        await context.close()
        memory[key] = session
        main = Math.max(main, session.summary.mainOwnedBytes)
        worker = Math.max(worker, session.summary.workerOwnedBytes)
        bitmap = Math.max(bitmap, session.summary.bitmapPeakBytes)
      }
      raw.memory = memory
      measurements.memory = { mainOwnedBytes: Number.isFinite(main) ? main : null, workerOwnedBytes: worker || null, bitmapPeakBytes: bitmap || null }
    }
  } finally {
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
      startedAt, finishedAt: new Date().toISOString(), build: 'vite.perf.config.ts production bundle, served from loopback with COOP/COEP', playwright: 'channel-launched headless', ...env.extra,
    },
    fixtures: Object.values(fixtures).map(f => ({ id: f.key, file: f.path.split('/').pop()!, bytes: f.bytes, sha256: f.sha256, uncompressedBytes: f.uncompressedBytes, nnz: f.nnz, scale: f.scale })),
    measurements,
    raw,
  }
  const profileJson = JSON.parse(readFileSync(join(repoRoot, 'tests/fixtures/xlsx-viewer/performance-profile.json'), 'utf8'))
  const receipt = await finalizeReceipt(draft, profileJson)
  const out = resolve(args.output)
  mkdirSync(out, { recursive: true })
  const path = join(out, `viewer-perf-${args.browser}-${args.profile}-${args.mode}-${ctx.label.slice(0, 12)}-${startedAt.replace(/[:.]/g, '-')}.json`)
  writeFileSync(path, JSON.stringify(receipt, null, 2) + '\n')
  return { receipt, path }
}
