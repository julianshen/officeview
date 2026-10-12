/**
 * S10.2 real-browser performance receipt (pure; shared by the perf page, the
 * bench CLI and unit tests).
 *
 * The receipt is hash-tagged like the S0.3/S8.4 results: SHA-256 over its
 * canonical JSON with `hash` removed. That detects accidental or casual edits;
 * it is not a signature or a machine attestation.
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
  [extra: string]: unknown
}
export interface ReceiptFixture { id: string; file: string; bytes: number; sha256: string; contentSha256?: string; uncompressedBytes: number; nnz: number; rows?: number; cols?: number; scale: string; route?: string }
export type ScrollKind = 'ordinary' | 'wrapped' | 'styled'
export type AckOp = 'sort' | 'filter' | 'find' | 'cancel'
export type ViewportFixture = 'F1' | 'F2' | 'F3'
export interface ReceiptMeasurements {
  scroll?: Partial<Record<ScrollKind, { inputToPaint: SeriesSummary; frameInterval: SeriesSummary }>>
  ack?: Partial<Record<AckOp, SeriesSummary>>
  firstViewport?: Partial<Record<ViewportFixture, SeriesSummary>>
  import?: Partial<Record<ViewportFixture, { expandedMiBPerSec: number; cellsPerSec: number }>>
  memory?: { mainOwnedBytes: number | null; workerOwnedBytes: number | null; bitmapPeakBytes: number | null }
}
export interface ReceiptDraft {
  schema: 'officeview-viewer-perf/1'
  mode: 'smoke' | 'release'
  environment: ReceiptEnvironment
  fixtures: ReceiptFixture[]
  /** Threshold inputs (summaries); everything else lives in `raw`. */
  measurements: ReceiptMeasurements
  /** Raw samples, phases, ledgers and notes, retained verbatim. */
  raw: Record<string, unknown>
}
export type Outcome = 'pass' | 'pass-limited' | 'miss' | 'not-run'
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
  evaluation: EvaluationRow[]
  /** Desktop release admission: no miss and no not-run desktop row. */
  desktopAdmitted: boolean
  hash: string
}

interface ProfileThreshold {
  id: string
  desktop: { frozenMs: number; outcome: string; measurement: { scenario: string; summary: { p95: number; samples: number } } | null }
  mobile: { frozenMs: number }
}
interface PerformanceProfile { thresholds: ProfileThreshold[]; evidence?: { commit?: string; results?: string } }

/** Spec §12.3/§12.4 values (not restated in the frozen profile). */
export const SPEC_BUDGETS = {
  'import-expanded-throughput': { kind: 'floor', unit: 'MiB/s', desktop: 20, mobile: 8, spec: '§12.3' },
  'import-cell-throughput': { kind: 'floor', unit: 'cells/s', desktop: 100_000, mobile: 30_000, spec: '§12.3' },
  'main-owned-memory': { kind: 'ceiling', unit: 'bytes', desktop: 64 * MiB, mobile: 32 * MiB, spec: '§12.4' },
  'worker-owned-memory': { kind: 'ceiling', unit: 'bytes', desktop: 512 * MiB, mobile: 256 * MiB, spec: '§12.4' },
  'bitmap-retained': { kind: 'ceiling', unit: 'bytes', desktop: 128 * MiB, mobile: 64 * MiB, spec: '§12.4' },
} as const
const KERNEL_ROWS = new Set(['numeric-sort-100k', 'numeric-sort-1m', 'numeric-filter-1m', 'natural-text-sort-1m', 'cold-value-index-100k'])

type Gathered = { value: number | null; limited: boolean; missing: string[]; note: string }
/** Worst value across required series; p95 when admitted, max when the sample count is limited. */
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
  return gathered.limited ? 'pass-limited' : 'pass'
}

/** Every threshold row for both profiles. Mobile rows are always not-run: no physical device run exists. */
export function evaluateReceipt(draft: ReceiptDraft, profile: PerformanceProfile): EvaluationRow[] {
  const m = draft.measurements
  const rows: EvaluationRow[] = []
  const measuredDesktop = draft.environment.profile === 'desktop'
  const push = (id: string, kind: 'ceiling' | 'floor', unit: string, desktopLimit: number, mobileLimit: number, gathered: Gathered, source: string) => {
    const outcome = measuredDesktop ? judge(kind, desktopLimit, gathered) : 'not-run'
    const missing = gathered.missing.length ? `missing: ${gathered.missing.join(', ')}` : ''
    rows.push({ id, profile: 'desktop', kind, limit: desktopLimit, unit, value: gathered.value, outcome, source, note: [gathered.note, missing].filter(Boolean).join('; ') })
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
        series[`${kind} input-to-paint`] = m.scroll?.[kind]?.inputToPaint
        series[`${kind} frame-interval`] = m.scroll?.[kind]?.frameInterval
      }
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling(series), 'S10.2 viewer sustained scroll (Chrome)')
    } else if (id === 'input-acknowledgment') {
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling({ sort: m.ack?.sort, filter: m.ack?.filter, find: m.ack?.find, cancel: m.ack?.cancel }), 'S10.2 viewer controls while the worker runs (Chrome)')
    } else if (id === 'progressive-first-viewport') {
      push(id, 'ceiling', 'ms', threshold.desktop.frozenMs, threshold.mobile.frozenMs, worstCeiling({ F1: m.firstViewport?.F1, F2: m.firstViewport?.F2, F3: m.firstViewport?.F3 }), 'S10.2 local File -> labelled usable viewport (Chrome)')
    }
  }
  const floor = (key: 'expandedMiBPerSec' | 'cellsPerSec'): Gathered => {
    let value: number | null = null
    const missing: string[] = [], parts: string[] = []
    for (const fixture of ['F1', 'F2', 'F3'] as const) {
      const v = m.import?.[fixture]?.[key]
      if (v === undefined) { missing.push(fixture); continue }
      parts.push(`${fixture} ${round(v)}`)
      value = value === null ? v : Math.min(value, v)
    }
    return { value, limited: false, missing, note: parts.join('; ') }
  }
  const budget = (key: keyof typeof SPEC_BUDGETS) => SPEC_BUDGETS[key]
  for (const [key, field] of [['import-expanded-throughput', 'expandedMiBPerSec'], ['import-cell-throughput', 'cellsPerSec']] as const) {
    const b = budget(key)
    push(key, b.kind, b.unit, b.desktop, b.mobile, floor(field), `S10.2 importer in a Chrome worker (${b.spec})`)
  }
  for (const [key, field] of [['main-owned-memory', 'mainOwnedBytes'], ['worker-owned-memory', 'workerOwnedBytes'], ['bitmap-retained', 'bitmapPeakBytes']] as const) {
    const b = budget(key)
    const v = m.memory?.[field] ?? null
    push(key, b.kind, b.unit, b.desktop, b.mobile, { value: v, limited: false, missing: v === null ? [field] : [], note: v === null ? '' : `${round(v / MiB)} MiB` }, `S10.2 viewer ledger/heap (${b.spec})`)
  }
  return rows
}

export async function finalizeReceipt(draft: ReceiptDraft, profile: PerformanceProfile): Promise<ViewerReceipt> {
  const evaluation = evaluateReceipt(draft, profile)
  const desktopAdmitted = evaluation.filter(r => r.profile === 'desktop').every(r => r.outcome === 'pass' || r.outcome === 'pass-limited')
  const body = { ...draft, evaluation, desktopAdmitted }
  return { ...body, hash: await sha256Hex(canonicalize(body)) }
}

export async function verifyReceipt(receipt: ViewerReceipt): Promise<{ ok: true } | { ok: false; reason: string }> {
  const { hash, ...body } = receipt
  const expected = await sha256Hex(canonicalize(body))
  return expected === hash ? { ok: true } : { ok: false, reason: `hash mismatch: recorded ${hash}, computed ${expected}` }
}
