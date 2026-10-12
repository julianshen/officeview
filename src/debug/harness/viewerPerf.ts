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
    /** Main ledger high-water mark (spec §12.4 peak including copies). */
    mainLedgerPeakBytes: number | null
    /** Largest live heap above the warmed-host baseline that the ledger peak does not cover. */
    mainUnattributedBytes: number | null
    /** Forced-GC probe noise of the warmed host (sample spread): the only allowed residual. */
    mainProbeNoiseBytes: number | null
    /** Largest amount by which a checkpoint's forced-GC heap above the warmed
     * host exceeds the ledger's CURRENT bytes there plus the allowance (the
     * warmed host's own open gap + noise); > 0 means steady uncharged memory. */
    mainCheckpointExcessBytes: number | null
    /** The warmed host's own gap, open and idle: forced-GC heap above the closed
     * warm host minus its ledger's current bytes (the smaller session's). */
    mainWarmOpenGapBytes: number | null
    /** Largest measured compiled-code growth over the warmed host at a gated
     * checkpoint (excluded from the gap; heap snapshot 'code' nodes). */
    mainCodeGrowthBytes: number | null
    /** Independent noise bound: max(idle probe noise, the checkpoint procedure's
     * spread on the idle closed warm host); the smaller session's. */
    mainIndependentNoiseBytes: number | null
    /** Largest warm-up residue of plain JS objects, arrays and native DOM. */
    mainWarmAppResidueBytes: number | null
    /** The explicit engine-cache allowance applied (user decision 2026-10-13). */
    mainEngineCacheAllowanceBytes: number | null
    /** Largest spread of a gated checkpoint's repeated forced-GC readings. */
    mainCheckpointSpreadBytes: number | null
    /** Largest excluded growth of the page's layout-shift entries (count x LAYOUT_SHIFT_ENTRY_BYTES). */
    mainLayoutShiftBytes: number | null
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

/** A scroll input with no present after it (the frame never showed it). */
export const SCROLL_CENSORED_MS = 1_000
/** Minimum forced-GC samples of the warmed host (baseline median and probe-noise spread). */
export const MIN_WARM_HOST_SAMPLES = 5
/** Bytes of compiled code in a V8 heap snapshot (nodes of type 'code':
 * bytecode, machine code, feedback and their metadata). */
export function snapshotCodeBytes(snapshot: unknown): number | null {
  const s = snapshot as { snapshot?: { meta?: { node_fields?: string[]; node_types?: unknown[] } }; nodes?: number[] }
  const fields = s.snapshot?.meta?.node_fields, types = s.snapshot?.meta?.node_types?.[0], nodes = s.nodes
  if (!Array.isArray(fields) || !Array.isArray(types) || !Array.isArray(nodes)) return null
  const width = fields.length, type = fields.indexOf('type'), size = fields.indexOf('self_size'), code = types.indexOf('code')
  if (type < 0 || size < 0 || code < 0 || width === 0) return null
  let bytes = 0
  for (let i = 0; i + width <= nodes.length; i += width) if (nodes[i + type] === code) bytes += nodes[i + size]
  return bytes
}
/** Heap-snapshot self bytes by node type ('code', 'string', 'object', 'native', ...). */
export function snapshotTypeBytes(snapshot: unknown): Record<string, number> {
  const s = snapshot as { snapshot?: { meta?: { node_fields?: string[]; node_types?: unknown[] } }; nodes?: number[] }
  const fields = s.snapshot?.meta?.node_fields, types = s.snapshot?.meta?.node_types?.[0], nodes = s.nodes
  const out: Record<string, number> = {}
  if (!Array.isArray(fields) || !Array.isArray(types) || !Array.isArray(nodes)) return out
  const width = fields.length, type = fields.indexOf('type'), size = fields.indexOf('self_size'), name = fields.indexOf('name')
  const strings = (snapshot as { strings?: string[] }).strings ?? []
  for (let i = 0; i + width <= nodes.length; i += width) {
    const t = String(types[nodes[i + type]])
    out[t] = (out[t] ?? 0) + nodes[i + size]
    // Native DOM nodes (elements, text, document fragments), detached or not.
    if (t === 'native' && /^(Detached )?(<|#text|#comment|#document)/.test(strings[nodes[i + name]] ?? '')) out['native:dom'] = (out['native:dom'] ?? 0) + nodes[i + size]
  }
  return out
}
/** Warm-up residue kinds the validator does not attribute to the engine:
 * plain JS objects, JS arrays and native DOM nodes (review r5, decision 6). */
export const WARM_APP_KINDS = ['object', 'array', 'native:dom'] as const
/** Minimum forced-GC samples of the warm document open and idle (review round 4). */
export const MIN_WARM_OPEN_SAMPLES = 3
/** Minimum repeated forced-GC readings at each gated checkpoint (its own noise). */
export const MIN_CHECKPOINT_READINGS = 5
/** The page's layout-shift buffer holds at most this many entries (spec cap). */
export const LAYOUT_SHIFT_BUFFER = 150
/** A layout-shift entry carries at most this many attribution sources. */
export const LAYOUT_SHIFT_MAX_SOURCES = 5
/**
 * Per-structure lower bounds of the page's buffered layout-shift entries,
 * excluded only for measured counts (the recorder counts entries and their
 * sources). Self sizes from a Chrome 155 heap snapshot:
 * - LayoutShift 145 B per entry;
 * - LayoutShiftAttribution 38 B per source;
 * - two DOMRectReadOnly (previous/current) at 53 B each per source.
 * The buffer's hash-table growth is not excluded.
 */
export const LAYOUT_SHIFT_ENTRY_BYTES = 145
export const LAYOUT_SHIFT_SOURCE_BYTES = 38 + 2 * 53
/** Minimum rounds of the checkpoint procedure on the idle warm host. */
export const MIN_WARM_NOISE_ROUNDS = 3
/**
 * Explicit engine-cache allowance per gated checkpoint and for the warm-up
 * residue check: USER DECISION 2026-10-13. After every measured exclusion, the
 * remaining residuals (r15: F2 loaded 0.05 MiB, F2 closed 0.03 MiB, warm-up
 * arrays/objects 0.13 MiB) were traced to engine/page caches whose capacity
 * grows and does not shrink (V8 smi_string_cache, Blink IdToNodeMap) and to
 * lazily loaded module state, with an untraced remainder. The allowance is
 * fixed and recorded in every receipt (mainEngineCacheAllowanceBytes);
 * anything above it is still refused. The 64 MiB budget is unchanged.
 */
export const ENGINE_CACHE_ALLOWANCE_BYTES = 262_144
/** Recovery cycles: the mean of the last three post-close readings may exceed the first three by at most this. */
export const RECOVERY_GROWTH_BYTES = 2 * MiB
export const RECOVERY_CYCLES = 10
const strictlyIncreasing = (v: number[]) => v.every((x, i) => i === 0 || x > v[i - 1])
const nondecreasing = (v: number[]) => v.every((x, i) => i === 0 || x >= v[i - 1])

/** Every threshold input recomputed from the raw records; problems name what is missing or invalid. */
export function deriveMeasurements(raw: Raw, fixtures: Record<string, ManifestFixture>): { measurements: ReceiptMeasurements; problems: string[] } {
  const problems: string[] = []
  const m: ReceiptMeasurements = { scroll: {}, ack: {}, firstViewport: {}, import: {}, memory: { mainRetainedBytes: null, mainSampledLivePeakBytes: null, mainLedgerPeakBytes: null, mainUnattributedBytes: null, mainProbeNoiseBytes: null, mainCheckpointExcessBytes: null, mainIndependentNoiseBytes: null, mainWarmAppResidueBytes: null, mainEngineCacheAllowanceBytes: null, mainWarmOpenGapBytes: null, mainCodeGrowthBytes: null, mainCheckpointSpreadBytes: null, mainLayoutShiftBytes: null, workerOwnedBytes: null, bitmapPeakBytes: null } }

  const scroll = obj(raw.scroll)
  for (const kind of ['ordinary', 'wrapped', 'styled'] as const) {
    const run = obj(scroll?.[kind]), timeline = obj(run?.timeline)
    const frames = nums(timeline?.frames), events = nums(timeline?.events), presents = nums(timeline?.presents)
    if (!run || !frames || !events || !presents) { problems.push(`scroll.${kind}: no raw timeline`); continue }
    if (run.importComplete !== true) problems.push(`scroll.${kind}: the fixture import did not complete before measuring`)
    if (!strictlyIncreasing(frames)) { problems.push(`scroll.${kind}: frame timestamps are not strictly increasing`); continue }
    if (!nondecreasing(events) || !nondecreasing(presents) || [...events, ...presents].some(t => t < 0)) { problems.push(`scroll.${kind}: input/present timestamps are negative or out of order`); continue }
    const intervals = frames.slice(1).map((t, i) => t - frames[i])
    const paired = pairInputsToPresents(events, presents)
    // An input that never reached a present is a censored miss, never dropped.
    const latencies = [...paired.latencies, ...Array(paired.unmatched).fill(SCROLL_CENSORED_MS) as number[]]
    m.scroll[kind] = {
      inputToPaint: seriesSummary(`scroll.${kind}.inputToPaint`, latencies, MIN_FRAMES, problems),
      frameInterval: seriesSummary(`scroll.${kind}.frameInterval`, intervals, MIN_FRAMES, problems),
    }
  }

  const ack = obj(raw.ack)
  for (const [op, key] of [['sort', 'sort'], ['filter', 'filterApply'], ['find', 'find'], ['cancel', 'cancel']] as const) {
    const records = obj(ack?.[key])?.records
    if (!Array.isArray(records)) { problems.push(`ack.${op}: no raw records`); continue }
    const values = records.map((r, i) => {
      const rec = obj(r), input = num(rec?.inputAt), mutation = num(rec?.mutationAt), present = num(rec?.presentAt)
      // No input seen, or no qualifying presented change: censored at the probe window.
      if (input === null || input < 0 || present === null || mutation === null || !rec?.what) return ACK_CENSORED_MS
      if (!(input <= mutation && mutation <= present)) { problems.push(`ack.${op}[${i}]: input -> mutation -> present out of order`); return ACK_CENSORED_MS }
      return present - input
    })
    m.ack[op] = seriesSummary(`ack.${op}`, values, MIN_OPERATIONS, problems)
  }

  const fv = obj(raw.firstViewport)
  for (const key of ['F1', 'F2', 'F3'] as const) {
    const runs = obj(fv?.[key])?.runs
    if (!Array.isArray(runs)) { problems.push(`firstViewport.${key}: no raw runs`); continue }
    const usable = runs.map(r => num(obj(r)?.firstUsable))
    if (usable.some(v => v === null || v < 0)) problems.push(`firstViewport.${key}: a run never reached a valid usable viewport`)
    m.firstViewport[key] = seriesSummary(`firstViewport.${key}`, usable.filter((v): v is number => v !== null && v >= 0), MIN_OPERATIONS, problems)
  }

  const imp = obj(raw.import)
  for (const key of ['F1', 'F2', 'F3'] as const) {
    const runs = obj(imp?.[key])?.runs, expected = fixtures[key]
    if (!Array.isArray(runs)) { problems.push(`import.${key}: no raw runs`); continue }
    const ok: ImportThroughput[] = []
    runs.forEach((r, i) => {
      const run = obj(r), events = Array.isArray(run?.events) ? run!.events as ImportEvent[] : null
      const where = `import.${key}[${i}]`
      if (run?.outcome !== 'complete' || !events?.length) { problems.push(`${where}: did not complete`); return }
      const ts = events.map(e => e.t), bytes = events.map(e => e.expandedBytes), cells = events.map(e => e.populatedCells)
      if (![...ts, ...bytes, ...cells].every(v => typeof v === 'number' && Number.isFinite(v) && v >= 0) || !nondecreasing(ts) || !nondecreasing(bytes) || !nondecreasing(cells)) { problems.push(`${where}: progress is not finite, nonnegative and monotonic`); return }
      const done = events.at(-1)!
      if (done.phase !== 'done') { problems.push(`${where}: no terminal done event`); return }
      // The terminal workload must be the verified fixture's (spec §12.3/§12.5).
      if (expected && done.populatedCells !== expected.nnz) { problems.push(`${where}: ${done.populatedCells} populated cells, the fixture holds ${expected.nnz}`); return }
      if (expected && (done.expandedBytes > expected.uncompressedBytes || done.expandedBytes < expected.uncompressedBytes * 0.99)) { problems.push(`${where}: ${done.expandedBytes} expanded bytes, the fixture expands to ${expected.uncompressedBytes}`); return }
      const throughput = importThroughput(events)
      if (!throughput || throughput.windowMs <= 0) { problems.push(`${where}: no post-inventory window`); return }
      ok.push(throughput)
    })
    if (ok.length < MIN_IMPORT_RUNS) problems.push(`import.${key}: ${ok.length} valid complete runs, below ${MIN_IMPORT_RUNS}`)
    // Lower-bound targets: the worst run is the reported value.
    if (ok.length) m.import[key] = { expandedMiBPerSec: Math.min(...ok.map(t => t.expandedMiBPerSec)), cellsPerSec: Math.min(...ok.map(t => t.cellsPerSec)), runs: ok.length }
  }

  // Memory: both sessions must provide every pool's evidence at every checkpoint
  // and a clean close; a maximum is only taken once nothing is missing.
  const memory = obj(raw.memory)
  const perSession: { retained: number; live: number; ledger: number; unattributed: number; noise: number; worker: number; bitmap: number; excess: number | null; warmOpenGap: number; code: number; gateMeasured: boolean; spread: number; shifts: number; independentNoise: number; warmAppResidue: number | null }[] = []
  for (const key of ['F2', 'F4'] as const) {
    const where = `memory.${key}`
    const session = obj(memory?.[key]), checkpoints = obj(session?.checkpoints)
    if (!session || !checkpoints || session.stalled) { problems.push(`${where}: no completed session`); continue }
    const before = problems.length
    if (!Array.isArray(session.timeouts)) problems.push(`${where}: no probe record`)
    else if (session.timeouts.length) problems.push(`${where}: probes timed out or failed (${(session.timeouts as string[]).join(', ')})`)
    const base = num(obj(checkpoints.baseline)?.heapBytes)
    if (base === null) problems.push(`${where}: no baseline heap`)
    let retained = -Infinity, worker = -Infinity, bitmap = -Infinity, ledgerPeak = -Infinity
    for (const name of ['loaded', 'afterOps'] as const) {
      const cp = obj(checkpoints[name])
      if (!cp) { problems.push(`${where}.${name}: missing checkpoint`); continue }
      const heap = num(cp.heapBytes)
      if (heap === null || base === null) problems.push(`${where}.${name}: no heap reading`)
      else retained = Math.max(retained, heap - base)
      const ws = obj(cp.worker)
      const peaks = (ws?.kind === 'stats' && Array.isArray(ws.sessions) ? ws.sessions as Raw[] : []).map(x => num(obj(x.memory)?.peakBytes)).filter((v): v is number => v !== null)
      if (!peaks.length) problems.push(`${where}.${name}: no worker ledger`)
      else worker = Math.max(worker, ...peaks)
      const uasm = obj(cp.uasm)
      const uw = uasm?.available === true ? num(uasm.worker) : null, uc = uasm?.available === true ? num(uasm.canvas) : null
      if (uw === null || uc === null) problems.push(`${where}.${name}: no measureUserAgentSpecificMemory reading`)
      else { worker = Math.max(worker, uw); bitmap = Math.max(bitmap, uc) }
      const peakBitmap = num(obj(cp.frame)?.peakBitmapBytes)
      if (peakBitmap === null || peakBitmap <= 0) problems.push(`${where}.${name}: no bitmap ledger`)
      else bitmap = Math.max(bitmap, peakBitmap)
      const ledger = obj(cp.ledger), lp = num(ledger?.peakBytes)
      if (lp === null) problems.push(`${where}.${name}: no main ledger snapshot`)
      else ledgerPeak = Math.max(ledgerPeak, lp)
    }
    for (const poll of Array.isArray(session.polls) ? session.polls as Raw[] : []) { const peak = num(poll.workerPeak); if (peak !== null) worker = Math.max(worker, peak) }
    // Close / recovery leg (AC-LIFE-02).
    const closed = obj(checkpoints.closed)
    const closedHeap = num(closed?.heapBytes), frame = obj(closed?.frame), canvas = obj(closed?.canvas), closedLedger = obj(closed?.ledger)
    if (!closed) problems.push(`${where}.closed: missing checkpoint`)
    else {
      const created = num(frame?.workersCreated), terminated = num(frame?.workersTerminated)
      if (created === null || terminated === null || created < 1 || created !== terminated) problems.push(`${where}.closed: workers ${terminated}/${created} terminated`)
      if (num(canvas?.count) !== 0) problems.push(`${where}.closed: ${String(canvas?.count)} canvases remain`)
      if (num(closedLedger?.bytes) !== 0) problems.push(`${where}.closed: main ledger holds ${String(closedLedger?.bytes)} B after close`)
      const cu = obj(closed.uasm)
      if (cu?.available !== true || num(cu.worker) !== 0) problems.push(`${where}.closed: worker memory not shown released`)
      const storage = num(obj(closed.storage)?.usage), baseStorage = num(obj(obj(checkpoints.baseline)?.storage)?.usage)
      if (storage === null || baseStorage === null || storage > baseStorage) problems.push(`${where}.closed: page storage not back to baseline`)
      if (closedHeap === null) problems.push(`${where}.closed: no heap reading`)
    }
    const cycles = nums(obj(session.summary)?.recovery && obj(obj(session.summary)!.recovery)?.cyclesMainOwnedBytes)
    if (key === 'F2') {
      if (!cycles || cycles.length < RECOVERY_CYCLES) problems.push(`${where}: ${cycles?.length ?? 0} recovery cycles, below ${RECOVERY_CYCLES}`)
      else {
        const mean = (v: number[]) => v.reduce((x, y) => x + y, 0) / v.length
        const growth = mean(cycles.slice(-3)) - mean(cycles.slice(0, 3))
        if (growth > RECOVERY_GROWTH_BYTES) problems.push(`${where}: main heap grows ${round(growth / MiB)} MiB over ${cycles.length} open/close cycles`)
      }
    }
    const liveSamples = nums(session.liveHeapSamples)
    if (!liveSamples?.length) problems.push(`${where}: no forced-GC live samples`)
    const warm = nums(obj(session.warmHost)?.samples)
    if (!warm || warm.length < MIN_WARM_HOST_SAMPLES) problems.push(`${where}: ${warm?.length ?? 0} warmed-host samples, below ${MIN_WARM_HOST_SAMPLES}`)
    const warmOpenRaw = obj(session.warmHost)?.open
    const warmOpen = Array.isArray(warmOpenRaw) ? (warmOpenRaw as Raw[]).map(o => ({ heapBytes: num(o.heapBytes), ledgerBytes: num(o.ledgerBytes) })).filter((o): o is { heapBytes: number; ledgerBytes: number } => o.heapBytes !== null && o.ledgerBytes !== null && o.ledgerBytes >= 0) : null
    if (!warmOpen || warmOpen.length < MIN_WARM_OPEN_SAMPLES) problems.push(`${where}: ${warmOpen?.length ?? 0} warm-open samples (heap + ledger), below ${MIN_WARM_OPEN_SAMPLES}`)
    const wh = obj(session.warmHost)
    if (problems.length > before || base === null || closedHeap === null || !liveSamples?.length || !warm || !warmOpen) continue
    // ---- Attribution (rounds 3-5). The baseline is the warmed host: the same
    // page with a small workbook opened, exercised and closed before the session.
    // Live peak: every forced-GC live sample above it must lie within the
    // ledger peak plus the idle probe noise.
    const sorted = [...warm].sort((x, y) => x - y)
    const warmBase = sorted[Math.floor(sorted.length / 2)]
    const noise = sorted[sorted.length - 1] - sorted[0]
    const live = Math.max(...liveSamples) - warmBase
    // Per checkpoint (loaded, afterOps, closed): forced-GC heap above the
    // warmed host minus the ledger's CURRENT bytes there, minus two measured,
    // page-level exclusions: compiled-code growth (heap-snapshot 'code' bytes)
    // and growth of the page's buffered layout-shift entries (per-structure
    // lower bounds). The allowance is the warmed host's own open gap plus
    // INDEPENDENT noise: the larger of the idle probe noise and the spread of
    // the checkpoint procedure repeated on the idle closed warm host. A
    // checkpoint's own reading spread can only tighten that noise term.
    // Every input is bounds-checked; anything missing or impossible leaves the
    // gate unmeasured (row pending), never silently excluded.
    let gateMeasured = true, warmAppResidue: number | null = null
    const block = (message: string) => { problems.push(`${where}: ${message}`); gateMeasured = false }
    const count = (v: unknown, max: number): number | null => { const n = num(v); return n !== null && Number.isInteger(n) && n >= 0 && n <= max ? n : null }
    const warmCode = num(wh?.codeBytes), warmOpenCode = num(wh?.openCodeBytes)
    const warmShifts = count(wh?.layoutShifts, LAYOUT_SHIFT_BUFFER), warmOpenShifts = count(wh?.openLayoutShifts, LAYOUT_SHIFT_BUFFER)
    const warmSources = count(wh?.layoutShiftSources, LAYOUT_SHIFT_BUFFER * LAYOUT_SHIFT_MAX_SOURCES), warmOpenSources = count(wh?.openLayoutShiftSources, LAYOUT_SHIFT_BUFFER * LAYOUT_SHIFT_MAX_SOURCES)
    const minOpenHeap = Math.min(...warmOpen.map(o => o.heapBytes))
    if (warmCode === null || warmOpenCode === null) block('no compiled-code measurement of the warmed host')
    else if (warmCode < 0 || warmCode > sorted[0] || warmOpenCode < 0 || warmOpenCode > minOpenHeap) block('warmed-host compiled-code bytes outside its measured heap')
    if (warmShifts === null || warmOpenShifts === null || warmSources === null || warmOpenSources === null) block(`no valid layout-shift entry/source counts of the warmed host (integers within the ${LAYOUT_SHIFT_BUFFER}-entry buffer)`)
    else if (warmSources > warmShifts * LAYOUT_SHIFT_MAX_SOURCES || warmOpenSources > warmOpenShifts * LAYOUT_SHIFT_MAX_SOURCES) block(`warmed-host layout-shift sources exceed ${LAYOUT_SHIFT_MAX_SOURCES} per entry`)
    const noiseRounds = Array.isArray(wh?.checkpointNoise) ? (wh!.checkpointNoise as unknown[]).map(r => nums(r)).filter((r): r is number[] => !!r && r.length >= MIN_CHECKPOINT_READINGS && r.every(v => v >= 0)) : []
    if (noiseRounds.length < MIN_WARM_NOISE_ROUNDS) block(`${noiseRounds.length} rounds of the checkpoint procedure on the idle warm host, below ${MIN_WARM_NOISE_ROUNDS}`)
    const procedureNoise = noiseRounds.length ? Math.max(...noiseRounds.map(r => Math.max(...r) - Math.min(...r))) : 0
    const independentNoise = Math.max(noise, procedureNoise)
    // Warm-up residue (decision 6): the closed warm host minus the empty page,
    // by heap-snapshot kind. Engine/code kinds (compiled code, script sources,
    // binding templates, hidden classes, engine caches) are allowed; plain JS
    // objects, arrays and native DOM above the independent noise are an
    // unattributed warm-up allocation (it would sit in the baseline).
    const emptyTypes = obj(wh?.emptyTypeBytes), warmTypes = obj(wh?.typeBytes)
    if (!emptyTypes || !warmTypes) block('no heap-snapshot type split of the empty page and the warm host')
    else {
      const residue = WARM_APP_KINDS.reduce((sum, kind) => sum + ((num(warmTypes[kind]) ?? 0) - (num(emptyTypes[kind]) ?? 0)), 0)
      warmAppResidue = residue
      if (residue > independentNoise + ENGINE_CACHE_ALLOWANCE_BYTES) block(`unattributed warm-up allocation: ${round(residue / MiB)} MiB of plain JS objects, arrays and native DOM in the warm host above the empty page (allowance: independent noise ${round(independentNoise / MiB)} MiB + engine-cache allowance ${round(ENGINE_CACHE_ALLOWANCE_BYTES / MiB)} MiB, user decision 2026-10-13)`)
    }
    const codeDelta = (bytes: number) => Math.max(0, bytes - (warmCode ?? bytes))
    const shiftDelta = (entries: number, sources: number) => Math.max(0, entries - (warmShifts ?? entries)) * LAYOUT_SHIFT_ENTRY_BYTES + Math.max(0, sources - (warmSources ?? sources)) * LAYOUT_SHIFT_SOURCE_BYTES
    const openGaps = warmOpen.map(o => o.heapBytes - o.ledgerBytes - warmBase - codeDelta(warmOpenCode ?? 0) - shiftDelta(warmOpenShifts ?? 0, warmOpenSources ?? 0)).sort((x, y) => x - y)
    const warmOpenGap = openGaps[Math.floor(openGaps.length / 2)]
    const openAllowance = Math.max(0, warmOpenGap)
    let excess: number | null = null, codeMax = 0, spreadMax = 0, shiftMax = 0
    for (const name of ['loaded', 'afterOps', 'closed'] as const) {
      const cp = obj(checkpoints[name]), current = num(obj(cp?.ledger)?.bytes), code = num(cp?.codeBytes)
      const readings = nums(cp?.heapReadings)
      const shifts = count(cp?.layoutShifts, LAYOUT_SHIFT_BUFFER), sources = count(cp?.layoutShiftSources, LAYOUT_SHIFT_BUFFER * LAYOUT_SHIFT_MAX_SOURCES)
      if (!readings || readings.length < MIN_CHECKPOINT_READINGS || readings.some(v => v < 0)) { block(`${name}: ${readings?.length ?? 0} valid repeated heap readings, below ${MIN_CHECKPOINT_READINGS}`); continue }
      const ordered = [...readings].sort((x, y) => x - y)
      const heap = ordered[Math.floor(ordered.length / 2)], spread = ordered[ordered.length - 1] - ordered[0]
      if (num(cp?.heapBytes) !== heap) { block(`${name}: recorded heap reading is not the median of its repeated readings`); continue }
      if (code === null) { block(`${name}: no compiled-code measurement`); continue }
      if (code < 0 || code > ordered[0]) { block(`${name}: compiled-code bytes outside the measured heap`); continue }
      if (shifts === null || sources === null) { block(`${name}: no valid layout-shift entry/source counts (integers within the ${LAYOUT_SHIFT_BUFFER}-entry buffer)`); continue }
      if (sources > shifts * LAYOUT_SHIFT_MAX_SOURCES) { block(`${name}: layout-shift sources exceed ${LAYOUT_SHIFT_MAX_SOURCES} per entry`); continue }
      if (current === null) { problems.push(`${where}.${name}: no ledger reading`); continue }
      const noiseTerm = Math.min(independentNoise, spread)
      // + the explicit engine-cache allowance (user decision 2026-10-13).
      const allowance = openAllowance + noiseTerm + ENGINE_CACHE_ALLOWANCE_BYTES
      spreadMax = Math.max(spreadMax, spread)
      const delta = codeDelta(code), shiftBytes = shiftDelta(shifts, sources)
      codeMax = Math.max(codeMax, delta)
      shiftMax = Math.max(shiftMax, shiftBytes)
      const gap = heap - warmBase - current - delta - shiftBytes
      excess = Math.max(excess ?? -Infinity, gap - allowance)
      if (gap > allowance) problems.push(`main-owned-memory: uncharged main memory in ${key} at the ${name} checkpoint (${round(gap / MiB)} MiB of forced-GC heap above the warmed host is not in the ledger's current ${round(current / MiB)} MiB after excluding ${round(delta / MiB)} MiB of measured compiled-code growth and ${round(shiftBytes / MiB)} MiB of layout-shift entries; allowance: the warmed host's open gap ${round(openAllowance / MiB)} MiB + independent noise ${round(noiseTerm / MiB)} MiB (idle probe ${round(noise / MiB)} MiB, warm-host checkpoint procedure ${round(procedureNoise / MiB)} MiB, capped by this checkpoint's spread ${round(spread / MiB)} MiB) + engine-cache allowance ${round(ENGINE_CACHE_ALLOWANCE_BYTES / MiB)} MiB (user decision 2026-10-13))`)
    }
    if (excess === null) gateMeasured = false
    perSession.push({ retained, live, ledger: ledgerPeak, unattributed: live - ledgerPeak, noise, worker, bitmap, excess, warmOpenGap, code: codeMax, gateMeasured, spread: spreadMax, shifts: shiftMax, independentNoise, warmAppResidue })
  }
  if (perSession.length === 2) {
    const max = (f: (s: typeof perSession[number]) => number) => Math.max(...perSession.map(f))
    m.memory = { mainRetainedBytes: max(s => s.retained), mainSampledLivePeakBytes: max(s => s.live), mainLedgerPeakBytes: max(s => s.ledger), mainUnattributedBytes: max(s => s.unattributed), mainProbeNoiseBytes: Math.min(...perSession.map(s => s.noise)), mainCheckpointExcessBytes: perSession.every(s => s.gateMeasured && s.excess !== null) ? Math.max(...perSession.map(s => s.excess as number)) : null, mainIndependentNoiseBytes: Math.min(...perSession.map(s => s.independentNoise)), mainWarmAppResidueBytes: perSession.every(s => s.warmAppResidue !== null) ? Math.max(...perSession.map(s => s.warmAppResidue as number)) : null, mainEngineCacheAllowanceBytes: ENGINE_CACHE_ALLOWANCE_BYTES, mainWarmOpenGapBytes: Math.min(...perSession.map(s => s.warmOpenGap)), mainCodeGrowthBytes: perSession.every(s => s.gateMeasured) ? max(s => s.code) : null, mainCheckpointSpreadBytes: max(s => s.spread), mainLayoutShiftBytes: perSession.every(s => s.gateMeasured) ? max(s => s.shifts) : null, workerOwnedBytes: max(s => s.worker), bitmapPeakBytes: max(s => s.bitmap) }
    for (const [i, session] of perSession.entries()) {
      if (session.unattributed > session.noise) problems.push(`main-owned-memory: unattributed main memory in ${['F2', 'F4'][i]} (${round(session.unattributed / MiB)} MiB of live heap above the warmed host is not in the ledger peak; measured probe noise ${round(session.noise / MiB)} MiB)`)
    }
  }
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
    const mib = (v: number | null) => (v === null ? 'n/a' : `${round(v / MiB)} MiB`)
    // Admitted on the ledger peak, only when the ledger accounts for the live
    // heap; a sampled live heap above the budget is a demonstrated miss.
    const unattributed = (mem.mainUnattributedBytes !== null && mem.mainProbeNoiseBytes !== null && mem.mainUnattributedBytes > mem.mainProbeNoiseBytes)
      || (mem.mainCheckpointExcessBytes !== null && mem.mainCheckpointExcessBytes > 0)
      // Never silently excluded: without a measured code growth the gate cannot run.
      || (mem.mainLedgerPeakBytes !== null && mem.mainCodeGrowthBytes === null)
    const value = mem.mainLedgerPeakBytes === null ? null : Math.max(mem.mainLedgerPeakBytes, mem.mainSampledLivePeakBytes !== null && mem.mainSampledLivePeakBytes > b.desktop ? mem.mainSampledLivePeakBytes : 0)
    const note = `ledger peak ${mib(mem.mainLedgerPeakBytes)}; forced-GC live peak above the warmed host ${mib(mem.mainSampledLivePeakBytes)}; unattributed ${mib(mem.mainUnattributedBytes)} (allowed: measured probe noise ${mib(mem.mainProbeNoiseBytes)}); checkpoint heap over the ledger's current bytes (less measured code growth ${mib(mem.mainCodeGrowthBytes)}) beyond the warmed host's open gap ${mib(mem.mainWarmOpenGapBytes)}, independent noise ${mib(mem.mainIndependentNoiseBytes)} and engine-cache allowance ${mib(mem.mainEngineCacheAllowanceBytes)} (user decision 2026-10-13): ${mib(mem.mainCheckpointExcessBytes)}; warm-up objects/arrays/DOM ${mib(mem.mainWarmAppResidueBytes)}; retained ${mib(mem.mainRetainedBytes)}`
    push('main-owned-memory', b.kind, b.unit, b.desktop, b.mobile,
      { value, limited: false, missing: value === null ? ['main ledger peak'] : [], note, ...(unattributed ? { pending: 'pending: unattributed main memory' } : {}) },
      `S10.2 main-thread application ledger, cross-checked with forced-GC heap samples (${b.spec})`)
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
  // Benign diagnostics are excluded by name at capture time; anything recorded here is a run error.
  if (!Array.isArray(draft.raw?.errors)) problems.push('run error record missing')
  else for (const error of draft.raw.errors as unknown[]) problems.push(`run error: ${String(error)}`)
  return problems
}

// ---- receipt ---------------------------------------------------------------------

function assemble(draft: ReceiptDraft, ctx: EvidenceContext) {
  const derived = deriveMeasurements(draft.raw ?? {}, ctx.fixtures)
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
