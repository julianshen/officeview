/**
 * S10.2 real-browser performance receipt (pure; shared by the perf page, the
 * bench CLI and unit tests).
 *
 * The receipt is hash-tagged like the S0.3/S8.4 results: SHA-256 over its
 * canonical JSON with `hash` removed. That detects accidental or casual edits;
 * it is not a signature or a machine attestation.
 *
 * The hash is integrity only. Evidence validity is separate: every threshold
 * input is DERIVED from the raw records (timelines, acknowledgment records,
 * import progress events, memory checkpoints), never taken from stored
 * summaries; and admission additionally requires release identity (mode,
 * browser, channel, profile, viewport, build, commit, clean tree), the fixture
 * identities of the committed manifest, every required scenario having
 * succeeded, and the minimum sample counts (spec §12.6) with only the
 * documented long-operation exceptions.
 *
 * Threshold sources: the frozen time ceilings come from
 * tests/fixtures/xlsx-viewer/performance-profile.json (spec §12.1 freeze);
 * importer floors and resident pool budgets are the spec's own values
 * (§12.3, §12.4), which the profile does not restate. Kernel sort/filter/index
 * rows are not re-run here: they cite the S8.6 evidence frozen in the profile.
 */
import { canonicalize, percentile, sha256Hex } from './result'

const MiB = 1024 * 1024

// ---- frames -------------------------------------------------------------

export interface PresentCounters { basePaints: number; tileHits: number; prefetched: number; overlayPaints: number }

/**
 * Whether an onFrame report is a base-layer present. A present touches every
 * visible tile (a cache hit or a fresh paint); a prefetch batch paints only
 * off-screen tiles (basePaints and prefetched rise together); an overlay-only
 * report changes neither.
 */
export function classifyPresent(previous: PresentCounters | null, next: PresentCounters): boolean {
  const before = previous ?? { basePaints: 0, tileHits: 0, prefetched: 0, overlayPaints: 0 }
  const hits = next.tileHits - before.tileHits
  const visiblePaints = (next.basePaints - before.basePaints) - (next.prefetched - before.prefetched)
  return hits + visiblePaints > 0
}

/**
 * Each input (sorted ms) is paired with the first present at or after it. An
 * input whose present also serves the next input is coalesced: the viewer
 * produced one frame for both, which is legal, and its latency still counts.
 */
export function pairInputsToPresents(inputs: readonly number[], presents: readonly number[]): { latencies: number[]; coalesced: number; unmatched: number } {
  const latencies: number[] = []
  let coalesced = 0, unmatched = 0, p = 0
  for (let i = 0; i < inputs.length; i++) {
    while (p < presents.length && presents[p] < inputs[i]) p++
    if (p === presents.length) { unmatched++; continue }
    latencies.push(presents[p] - inputs[i])
    if (i + 1 < inputs.length && presents[p] >= inputs[i + 1]) coalesced++
  }
  return { latencies, coalesced, unmatched }
}

// ---- summaries ------------------------------------------------------------

export type SeriesSummary =
  | { status: 'ok'; samples: number; p50: number; p95: number; max: number; min: number; mean: number }
  | { status: 'insufficient-samples'; samples: number; p50: number; p95: null; max: number; min: number; mean: number }
  | { status: 'no-samples'; samples: 0 }

/** Nearest-rank p50/p95 (spec §12.6); p95 only at or above `minSamples`. */
export function summarize(values: readonly number[], minSamples: number): SeriesSummary {
  if (values.length === 0) return { status: 'no-samples', samples: 0 }
  const sorted = [...values].sort((a, b) => a - b)
  const mean = sorted.reduce((sum, v) => sum + v, 0) / sorted.length
  const base = { samples: sorted.length, p50: percentile(sorted, 50), max: sorted[sorted.length - 1], min: sorted[0], mean }
  return sorted.length >= minSamples ? { status: 'ok', ...base, p95: percentile(sorted, 95) } : { status: 'insufficient-samples', ...base, p95: null }
}

// ---- importer throughput ------------------------------------------------------

export interface ImportEvent { t: number; phase: 'inventory' | 'styles' | 'sharedStrings' | 'sheet' | 'done'; expandedBytes: number; populatedCells: number; compressedRead: number }
export interface ImportThroughput {
  /** Inventory event -> done: every expanded byte and cell after ZIP inventory/startup. */
  windowMs: number
  expandedBytes: number
  populatedCells: number
  expandedMiBPerSec: number
  cellsPerSec: number
  /** First committed sheet page -> done (diagnostic: excludes sharedStrings/styles). */
  sheetSlopeMiBPerSec: number | null
  sheetSlopeCellsPerSec: number | null
}

/**
 * Post-inventory importer throughput (spec §12.3: after ZIP inventory/startup,
 * bytes and cells reported separately). Null without an inventory and a done.
 */
export function importThroughput(events: readonly ImportEvent[]): ImportThroughput | null {
  const inventory = events.find(e => e.phase === 'inventory')
  const done = events.find(e => e.phase === 'done')
  if (!inventory || !done || done.t <= inventory.t) return null
  const windowMs = done.t - inventory.t
  const expandedBytes = done.expandedBytes - inventory.expandedBytes
  const populatedCells = done.populatedCells - inventory.populatedCells
  const firstSheet = events.find(e => e.phase === 'sheet')
  const sheetMs = firstSheet ? done.t - firstSheet.t : 0
  return {
    windowMs, expandedBytes, populatedCells,
    expandedMiBPerSec: expandedBytes / MiB / (windowMs / 1000),
    cellsPerSec: populatedCells / (windowMs / 1000),
    sheetSlopeMiBPerSec: firstSheet && sheetMs > 0 ? (done.expandedBytes - firstSheet.expandedBytes) / MiB / (sheetMs / 1000) : null,
    sheetSlopeCellsPerSec: firstSheet && sheetMs > 0 ? (done.populatedCells - firstSheet.populatedCells) / (sheetMs / 1000) : null,
  }
}

// ---- receipt ----------------------------------------------------------------

export interface ReceiptEnvironment {
  commit: string
  dirty: boolean
  browser: string
  browserVersion: string
  channel: string
  headless: boolean
  profile: 'desktop' | 'mobile'
  viewport: { width: number; height: number }
  dpr: number
  machine: string
  os: string
  loadAverage: number[]
  /** Free-text machine-contention note recorded with every run. */
  contention: string
  startedAt: string
  finishedAt: string
  build?: string
  [extra: string]: unknown
}
export interface ReceiptFixture { id: string; file: string; bytes: number; sha256: string; contentSha256?: string; uncompressedBytes: number; nnz: number; rows?: number; cols?: number; scale: string; route?: string }
export type ScrollKind = 'ordinary' | 'wrapped' | 'styled'
export type AckOp = 'sort' | 'filter' | 'find' | 'cancel'
export type ViewportFixture = 'F1' | 'F2' | 'F3'
export interface ReceiptMeasurements {
  scroll: Partial<Record<ScrollKind, { inputToPaint: SeriesSummary; frameInterval: SeriesSummary }>>
  ack: Partial<Record<AckOp, SeriesSummary>>
  firstViewport: Partial<Record<ViewportFixture, SeriesSummary>>
  import: Partial<Record<ViewportFixture, { expandedMiBPerSec: number; cellsPerSec: number; runs: number }>>
  memory: {
    /** Post-GC retained main heap owned by the viewer (worst checkpoint). */
    mainRetainedBytes: number | null
    /** Forced-GC live samples during the operations: a LOWER bound of the live peak. */
    mainSampledLivePeakBytes: number | null
    /** Application ledger high-water mark of main-owned live allocations; null: no complete ledger exists. */
    mainLedgerPeakBytes: number | null
    workerOwnedBytes: number | null
    bitmapPeakBytes: number | null
  }
}
export interface ReceiptDraft {
  schema: 'officeview-viewer-perf/1'
  mode: 'smoke' | 'release'
  environment: ReceiptEnvironment
  fixtures: ReceiptFixture[]
  /** Raw records: the only source of threshold inputs. */
  raw: Record<string, unknown>
}
/** pending: measured but the required evidence (e.g. a live peak) is not demonstrated. */
export type Outcome = 'pass' | 'pass-limited' | 'miss' | 'pending' | 'not-run'
export interface EvaluationRow {
  id: string
  profile: 'desktop' | 'mobile'
  kind: 'ceiling' | 'floor'
  limit: number
  unit: string
  value: number | null
  outcome: Outcome
  source: string
  note: string
}
export interface ViewerReceipt extends ReceiptDraft {
  /** Derived from `raw` (recomputed and compared by validateReceipt). */
  measurements: ReceiptMeasurements
  evaluation: EvaluationRow[]
  /** Why the evidence cannot admit a release (empty when it can). */
  evidenceProblems: string[]
  /** Release admission: valid evidence and every desktop row pass / pass-limited. */
  desktopAdmitted: boolean
  hash: string
}

interface ProfileThreshold {
  id: string
  desktop: { frozenMs: number; outcome: string; measurement: { scenario: string; summary: { p95: number; samples: number } } | null }
  mobile: { frozenMs: number }
}
export interface PerformanceProfile { thresholds: ProfileThreshold[]; evidence?: { commit?: string; results?: string } }
export interface ManifestFixture { contentSha256: string; uncompressedBytes: number; nnz: number; rows: number; cols: number }
export interface EvidenceContext {
  profile: PerformanceProfile
  /** Full-scale entries of tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json. */
  fixtures: Record<string, ManifestFixture>
}

/** Spec §12.3/§12.4 values (not restated in the frozen profile). */
export const SPEC_BUDGETS = {
  'import-expanded-throughput': { kind: 'floor', unit: 'MiB/s', desktop: 20, mobile: 8, spec: '§12.3' },
  'import-cell-throughput': { kind: 'floor', unit: 'cells/s', desktop: 100_000, mobile: 30_000, spec: '§12.3' },
  'main-owned-memory': { kind: 'ceiling', unit: 'bytes', desktop: 64 * MiB, mobile: 32 * MiB, spec: '§12.4' },
  'worker-owned-memory': { kind: 'ceiling', unit: 'bytes', desktop: 512 * MiB, mobile: 256 * MiB, spec: '§12.4' },
  'bitmap-retained': { kind: 'ceiling', unit: 'bytes', desktop: 128 * MiB, mobile: 64 * MiB, spec: '§12.4' },
} as const
const KERNEL_ROWS = new Set(['numeric-sort-100k', 'numeric-sort-1m', 'numeric-filter-1m', 'natural-text-sort-1m', 'cold-value-index-100k'])

/** Release identity of an admitted desktop D receipt. */
export const RELEASE_IDENTITY = { mode: 'release', browser: 'chrome', channel: 'chrome', profile: 'desktop', viewport: { width: 1440, height: 900 }, dpr: 2,
  build: 'vite.perf.config.ts production bundle, served from loopback with COOP/COEP' } as const
export const REQUIRED_FIXTURES = ['F1', 'F2', 'F3', 'F4', 'wrapped', 'styled'] as const
/** Spec §12.6 minimums, and the only documented long-operation exceptions. */
export const MIN_FRAMES = 300
export const MIN_OPERATIONS = 30
export const MIN_IMPORT_RUNS = 3
export const LIMITED_SERIES: Readonly<Record<string, number>> = Object.freeze({
  // 1M-row cold opens (each opens, imports and closes a fresh viewer).
  'firstViewport.F2': 10,
  // Each apply + clear re-runs a 1M-row value filter in the worker.
  'ack.filter': 10,
})
/** A censored acknowledgment (no qualifying presented change within the probe window). */
export const ACK_CENSORED_MS = 15_000

// ---- derivation from raw records --------------------------------------------

type Raw = Record<string, unknown>
const obj = (v: unknown): Raw | null => (v && typeof v === 'object' && !Array.isArray(v) ? v as Raw : null)
const nums = (v: unknown): number[] | null => (Array.isArray(v) && v.every(x => typeof x === 'number' && Number.isFinite(x)) ? v as number[] : null)
const num = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : null)

function minimumFor(series: string, standard: number): number { return LIMITED_SERIES[series] ?? standard }
function seriesSummary(series: string, values: number[], standard: number, problems: string[]): SeriesSummary {
  const minimum = minimumFor(series, standard)
  if (values.length < minimum) problems.push(`${series}: ${values.length} samples, below the minimum ${minimum}`)
  // p95 needs the full §12.6 count; a documented exception is judged on max.
  return summarize(values, standard)
}

/** Every threshold input recomputed from the raw records; problems name what is missing or invalid. */
export function deriveMeasurements(raw: Raw): { measurements: ReceiptMeasurements; problems: string[] } {
  const problems: string[] = []
  const m: ReceiptMeasurements = { scroll: {}, ack: {}, firstViewport: {}, import: {}, memory: { mainRetainedBytes: null, mainSampledLivePeakBytes: null, mainLedgerPeakBytes: null, workerOwnedBytes: null, bitmapPeakBytes: null } }

  const scroll = obj(raw.scroll)
  for (const kind of ['ordinary', 'wrapped', 'styled'] as const) {
    const run = obj(scroll?.[kind]), timeline = obj(run?.timeline)
    const frames = nums(timeline?.frames), events = nums(timeline?.events), presents = nums(timeline?.presents)
    if (!run || !frames || !events || !presents) { problems.push(`scroll.${kind}: no raw timeline`); continue }
    if (run.importComplete !== true) problems.push(`scroll.${kind}: the fixture import did not complete before measuring`)
    const intervals = frames.slice(1).map((t, i) => t - frames[i])
    const paired = pairInputsToPresents(events, presents)
    m.scroll[kind] = {
      inputToPaint: seriesSummary(`scroll.${kind}.inputToPaint`, paired.latencies, MIN_FRAMES, problems),
      frameInterval: seriesSummary(`scroll.${kind}.frameInterval`, intervals, MIN_FRAMES, problems),
    }
  }

  const ack = obj(raw.ack)
  for (const [op, key] of [['sort', 'sort'], ['filter', 'filterApply'], ['find', 'find'], ['cancel', 'cancel']] as const) {
    const records = obj(ack?.[key])?.records
    if (!Array.isArray(records)) { problems.push(`ack.${op}: no raw records`); continue }
    const values = records.map(r => {
      const rec = obj(r), input = num(rec?.inputAt), present = num(rec?.presentAt)
      // No input seen, or no qualifying presented change: censored at the probe window.
      return input !== null && input >= 0 && present !== null && rec?.what ? present - input : ACK_CENSORED_MS
    })
    m.ack[op] = seriesSummary(`ack.${op}`, values, MIN_OPERATIONS, problems)
  }

  const fv = obj(raw.firstViewport)
  for (const key of ['F1', 'F2', 'F3'] as const) {
    const runs = obj(fv?.[key])?.runs
    if (!Array.isArray(runs)) { problems.push(`firstViewport.${key}: no raw runs`); continue }
    const usable = runs.map(r => num(obj(r)?.firstUsable))
    if (usable.some(v => v === null)) problems.push(`firstViewport.${key}: a run never reached a usable viewport`)
    m.firstViewport[key] = seriesSummary(`firstViewport.${key}`, usable.filter((v): v is number => v !== null), MIN_OPERATIONS, problems)
  }

  const imp = obj(raw.import)
  for (const key of ['F1', 'F2', 'F3'] as const) {
    const runs = obj(imp?.[key])?.runs
    if (!Array.isArray(runs)) { problems.push(`import.${key}: no raw runs`); continue }
    const results = runs.map(r => (obj(r)?.outcome === 'complete' && Array.isArray(obj(r)?.events) ? importThroughput(obj(r)!.events as ImportEvent[]) : null))
    if (results.some(t => t === null)) problems.push(`import.${key}: a run did not complete`)
    const ok = results.filter((t): t is ImportThroughput => t !== null)
    if (ok.length < MIN_IMPORT_RUNS) problems.push(`import.${key}: ${ok.length} complete runs, below ${MIN_IMPORT_RUNS}`)
    // Lower-bound targets: the worst run is the reported value.
    if (ok.length) m.import[key] = { expandedMiBPerSec: Math.min(...ok.map(t => t.expandedMiBPerSec)), cellsPerSec: Math.min(...ok.map(t => t.cellsPerSec)), runs: ok.length }
  }

  const memory = obj(raw.memory)
  let retained: number | null = null, live: number | null = null, worker: number | null = null, bitmap: number | null = null
  for (const key of ['F2', 'F4'] as const) {
    const session = obj(memory?.[key]), checkpoints = obj(session?.checkpoints)
    if (!session || !checkpoints || session.stalled) { problems.push(`memory.${key}: no completed session`); continue }
    if (Array.isArray(session.timeouts) && session.timeouts.length) problems.push(`memory.${key}: probes timed out (${(session.timeouts as string[]).join(', ')})`)
    const base = num(obj(checkpoints.baseline)?.heapBytes)
    if (base === null) { problems.push(`memory.${key}: no baseline heap`); continue }
    for (const name of ['loaded', 'afterOps']) {
      const cp = obj(checkpoints[name])
      const heap = num(cp?.heapBytes)
      if (heap === null) problems.push(`memory.${key}.${name}: no heap reading`)
      else retained = Math.max(retained ?? -Infinity, heap - base)
      const ws = obj(cp?.worker)
      const sessions = ws?.kind === 'stats' && Array.isArray(ws.sessions) ? ws.sessions as Raw[] : []
      for (const sess of sessions) { const peak = num(obj(sess.memory)?.peakBytes); if (peak !== null) worker = Math.max(worker ?? 0, peak) }
      const uasm = obj(cp?.uasm)
      if (uasm?.available === true) {
        const w = num(uasm.worker), c = num(uasm.canvas)
        if (w !== null) worker = Math.max(worker ?? 0, w)
        if (c !== null) bitmap = Math.max(bitmap ?? 0, c)
      }
      const peakBitmap = num(obj(cp?.frame)?.peakBitmapBytes)
      if (peakBitmap !== null) bitmap = Math.max(bitmap ?? 0, peakBitmap)
    }
    for (const poll of Array.isArray(session.polls) ? session.polls as Raw[] : []) { const peak = num(poll.workerPeak); if (peak !== null) worker = Math.max(worker ?? 0, peak) }
    const liveSamples = nums(session.liveHeapSamples)
    if (liveSamples && liveSamples.length) live = Math.max(live ?? -Infinity, ...liveSamples.map(v => v - base))
    else problems.push(`memory.${key}: no forced-GC live samples`)
  }
  m.memory = { mainRetainedBytes: retained, mainSampledLivePeakBytes: live, mainLedgerPeakBytes: null, workerOwnedBytes: worker, bitmapPeakBytes: bitmap }
  // §12.4 budgets the main pool's PEAK including copies. Without a complete
  // main-owned application ledger the peak is not demonstrated.
  if (m.memory.mainLedgerPeakBytes === null) problems.push('main-owned-memory: no complete main-owned application ledger; the live peak is not demonstrated (row pending)')
  return { measurements: m, problems }
}

// ---- evaluation --------------------------------------------------------------

type Gathered = { value: number | null; limited: boolean; missing: string[]; note: string; pending?: string }
/** Worst value across required series: p95 at the §12.6 count, max for a documented exception. */
function worstCeiling(series: Record<string, SeriesSummary | undefined>): Gathered {
  let value: number | null = null, limited = false
  const missing: string[] = [], parts: string[] = []
  for (const [name, summary] of Object.entries(series)) {
    if (!summary || summary.status === 'no-samples') { missing.push(name); continue }
    const v = summary.status === 'ok' ? summary.p95 : summary.max
    if (summary.status !== 'ok') limited = true
    parts.push(`${name} ${summary.status === 'ok' ? 'p95' : `max (n=${summary.samples})`} ${round(v)}`)
    value = value === null ? v : Math.max(value, v)
  }
  return { value, limited, missing, note: parts.join('; ') }
}
const round = (v: number) => Math.round(v * 100) / 100

function judge(kind: 'ceiling' | 'floor', limit: number, gathered: Gathered): Outcome {
  if (gathered.value === null) return 'not-run'
  // A measured exceedance is a miss even while other required series are missing.
  const ok = kind === 'ceiling' ? gathered.value <= limit : gathered.value >= limit
  if (!ok) return 'miss'
  if (gathered.missing.length > 0) return 'not-run'
  if (gathered.pending) return 'pending'
  return gathered.limited ? 'pass-limited' : 'pass'
}

/** Every threshold row for both profiles from derived measurements. Mobile rows are always not-run. */
export function evaluateMeasurements(m: ReceiptMeasurements, environmentProfile: 'desktop' | 'mobile', profile: PerformanceProfile): EvaluationRow[] {
  const rows: EvaluationRow[] = []
  const measuredDesktop = environmentProfile === 'desktop'
  const push = (id: string, kind: 'ceiling' | 'floor', unit: string, desktopLimit: number, mobileLimit: number, gathered: Gathered, source: string) => {
    const outcome = measuredDesktop ? judge(kind, desktopLimit, gathered) : 'not-run'
    const missing = gathered.missing.length ? `missing: ${gathered.missing.join(', ')}` : ''
    rows.push({ id, profile: 'desktop', kind, limit: desktopLimit, unit, value: gathered.value, outcome, source, note: [gathered.note, missing, gathered.pending ?? ''].filter(Boolean).join('; ') })
    rows.push({ id, profile: 'mobile', kind, limit: mobileLimit, unit, value: null, outcome: 'not-run', source: 'no physical iPad/mobile run (skipped by user decision)', note: 'mobile ceiling frozen at the spec value; no device claim' })
  }
  for (const threshold of profile.thresholds) {
    const { id } = threshold
    if (KERNEL_ROWS.has(id)) {
      const measurement = threshold.desktop.measurement
      const value = measurement ? measurement.summary.p95 : null
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs,
        { value, limited: false, missing: measurement ? [] : ['S8.6 measurement'], note: measurement ? `${measurement.scenario} query-completion p95 (n=${measurement.summary.samples})` : '' },
        `S8.6 Chrome evidence reused (${profile.evidence?.results ?? 'performance-profile.json'} @ ${(profile.evidence?.commit ?? '').slice(0, 7)}); not re-run by S10.2`)
      continue
    }
    if (id === 'warm-viewport-frame') {
      const series: Record<string, SeriesSummary | undefined> = {}
      for (const kind of ['ordinary', 'wrapped', 'styled'] as const) {
        series[`${kind} input-to-paint`] = m.scroll[kind]?.inputToPaint
        series[`${kind} frame-interval`] = m.scroll[kind]?.frameInterval
      }
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling(series), 'S10.2 viewer sustained scroll (Chrome)')
    } else if (id === 'input-acknowledgment') {
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling({ sort: m.ack.sort, filter: m.ack.filter, find: m.ack.find, cancel: m.ack.cancel }), 'S10.2 viewer controls while the worker runs (Chrome)')
    } else if (id === 'progressive-first-viewport') {
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling({ F1: m.firstViewport.F1, F2: m.firstViewport.F2, F3: m.firstViewport.F3 }), 'S10.2 local File -> labelled usable viewport (Chrome)')
    }
  }
  const floor = (key: 'expandedMiBPerSec' | 'cellsPerSec'): Gathered => {
    let value: number | null = null
    const missing: string[] = [], parts: string[] = []
    for (const fixture of ['F1', 'F2', 'F3'] as const) {
      const v = m.import[fixture]?.[key]
      if (v === undefined) { missing.push(fixture); continue }
      parts.push(`${fixture} ${round(v)}`)
      value = value === null ? v : Math.min(value, v)
    }
    return { value, limited: false, missing, note: parts.join('; ') }
  }
  for (const [key, field] of [['import-expanded-throughput', 'expandedMiBPerSec'], ['import-cell-throughput', 'cellsPerSec']] as const) {
    const b = SPEC_BUDGETS[key]
    push(key, b.kind, b.unit, b.desktop, b.mobile, floor(field), `S10.2 importer in a Chrome worker (${b.spec})`)
  }
  {
    const b = SPEC_BUDGETS['main-owned-memory'], mem = m.memory
    // Admission needs the ledger peak; the sampled live peak can only prove a miss.
    const value = mem.mainLedgerPeakBytes ?? mem.mainSampledLivePeakBytes
    const note = `retained ${mem.mainRetainedBytes === null ? 'n/a' : `${round(mem.mainRetainedBytes / MiB)} MiB`}; forced-GC sampled live peak (lower bound) ${mem.mainSampledLivePeakBytes === null ? 'n/a' : `${round(mem.mainSampledLivePeakBytes / MiB)} MiB`}`
    push('main-owned-memory', b.kind, b.unit, b.desktop, b.mobile,
      { value, limited: false, missing: value === null ? ['main live peak'] : [], note, ...(mem.mainLedgerPeakBytes === null ? { pending: 'pending: no complete main-owned application ledger, so the §12.4 peak is not demonstrated' } : {}) },
      `S10.2 main heap probes (${b.spec})`)
  }
  for (const [key, field] of [['worker-owned-memory', 'workerOwnedBytes'], ['bitmap-retained', 'bitmapPeakBytes']] as const) {
    const b = SPEC_BUDGETS[key]
    const v = m.memory[field]
    push(key, b.kind, b.unit, b.desktop, b.mobile, { value: v, limited: false, missing: v === null ? [field] : [], note: v === null ? '' : `${round(v / MiB)} MiB` }, `S10.2 viewer ledger/heap (${b.spec})`)
  }
  return rows
}

// ---- identity ------------------------------------------------------------------

/** Why this receipt's identity cannot be release evidence (empty when it can). */
export function identityProblems(draft: ReceiptDraft, ctx: EvidenceContext): string[] {
  const problems: string[] = []
  const e = draft.environment, want = RELEASE_IDENTITY
  if (draft.schema !== 'officeview-viewer-perf/1') problems.push(`schema ${String(draft.schema)}`)
  if (draft.mode !== want.mode) problems.push(`mode ${draft.mode}, not release`)
  if (e.browser !== want.browser) problems.push(`browser ${e.browser}, not installed Chrome`)
  if (e.channel !== want.channel) problems.push(`channel ${e.channel}, not chrome`)
  if (!/^\d+\.\d+/.test(String(e.browserVersion ?? ''))) problems.push('browser version missing')
  if (e.profile !== want.profile) problems.push(`profile ${e.profile}, not desktop`)
  if (e.viewport?.width !== want.viewport.width || e.viewport?.height !== want.viewport.height) problems.push(`viewport ${e.viewport?.width}x${e.viewport?.height}, not 1440x900`)
  if (e.dpr !== want.dpr) problems.push(`DPR ${e.dpr}, not 2`)
  if (e.build !== want.build) problems.push(`build "${String(e.build)}", not the production perf bundle`)
  if (!/^[0-9a-f]{40}$/.test(String(e.commit))) problems.push(`commit "${String(e.commit)}" is not a full SHA`)
  if (e.dirty !== false) problems.push('the checkout was dirty')
  const byId = new Map((Array.isArray(draft.fixtures) ? draft.fixtures : []).map(f => [f.id, f]))
  for (const id of REQUIRED_FIXTURES) {
    const got = byId.get(id), expected = ctx.fixtures[id]
    if (!expected) { problems.push(`fixture manifest has no ${id}`); continue }
    if (!got) { problems.push(`fixture ${id} missing`); continue }
    for (const field of ['contentSha256', 'uncompressedBytes', 'nnz', 'rows', 'cols'] as const) {
      if (got[field] !== expected[field]) problems.push(`fixture ${id}: ${field} ${String(got[field])} != manifest ${String(expected[field])}`)
    }
  }
  const errors = Array.isArray(draft.raw?.errors) ? draft.raw.errors as unknown[] : []
  for (const error of errors) if (typeof error === 'string' && /disconnected|crashed|abandoned|timed out/.test(error)) problems.push(`run error: ${error}`)
  return problems
}

// ---- receipt ---------------------------------------------------------------------

function assemble(draft: ReceiptDraft, ctx: EvidenceContext) {
  const derived = deriveMeasurements(draft.raw ?? {})
  const evaluation = evaluateMeasurements(derived.measurements, draft.environment.profile, ctx.profile)
  const evidenceProblems = [...identityProblems(draft, ctx), ...derived.problems]
  const rowsOk = evaluation.filter(r => r.profile === 'desktop').every(r => r.outcome === 'pass' || r.outcome === 'pass-limited')
  return { measurements: derived.measurements, evaluation, evidenceProblems, desktopAdmitted: rowsOk && evidenceProblems.length === 0 }
}

export async function finalizeReceipt(draft: ReceiptDraft, ctx: EvidenceContext): Promise<ViewerReceipt> {
  const { schema, mode, environment, fixtures, raw } = draft
  const body = { schema, mode, environment, fixtures, raw, ...assemble(draft, ctx) }
  return { ...body, hash: await sha256Hex(canonicalize(body)) }
}

/** Integrity only: the hash over the canonical body. */
export async function verifyReceipt(receipt: ViewerReceipt): Promise<{ ok: true } | { ok: false; reason: string }> {
  const { hash, ...body } = receipt
  const expected = await sha256Hex(canonicalize(body))
  return expected === hash ? { ok: true } : { ok: false, reason: `hash mismatch: recorded ${hash}, computed ${expected}` }
}

/**
 * Evidence validation: integrity, then everything recomputed from the raw
 * records and compared with what the receipt states. A receipt admits only if
 * the recomputation admits it AND agrees with the stored rows and verdict.
 */
export async function validateReceipt(receipt: ViewerReceipt, ctx: EvidenceContext): Promise<{ admitted: boolean; problems: string[] }> {
  const problems: string[] = []
  const integrity = await verifyReceipt(receipt)
  if (!integrity.ok) problems.push(integrity.reason)
  const fresh = assemble(receipt, ctx)
  if (canonicalize(fresh.measurements) !== canonicalize(receipt.measurements ?? null)) problems.push('stored measurements disagree with the raw records')
  if (canonicalize(fresh.evaluation) !== canonicalize(receipt.evaluation ?? null)) problems.push('stored evaluation disagrees with the raw records')
  if (fresh.desktopAdmitted !== receipt.desktopAdmitted) problems.push(`stored verdict ${receipt.desktopAdmitted} disagrees with the recomputed ${fresh.desktopAdmitted}`)
  problems.push(...fresh.evidenceProblems)
  return { admitted: problems.length === 0 && fresh.desktopAdmitted, problems }
}
