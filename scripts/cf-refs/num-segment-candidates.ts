/**
 * The num-segment follow-up (brief 2026-11-22, Phase 1 task 1): a search for the arithmetic Excel 16.115 uses inside a
 * colour-scale segment, over every captured colour-scale cell.
 *
 * Capture e448d831 (stage2-scale-scope-disjoint twins T4 and V3: 3-colour literal num 0/30/60, values 40 and 50) drew one
 * unit higher per RGB channel than production's truncating interpolation. This script enumerates closed, deterministic
 * formulations of the segment colour and evaluates each on every captured colour-scale cell production paints exactly
 * today or withholds as the num-segment shape. A candidate SURVIVES a scope when it reproduces Excel's DisplayFormat fill
 * on every in-scope cell (T4 and V3 among them); cells outside the scope keep production's arithmetic, which already
 * equals Excel there. Survivors are hypotheses to pre-register for a NEW capture; nothing here is registered.
 *
 * How the cells are found (no re-derived statistics): production's own pipeline (`prepareCfPaint`, as the conformance test
 * runs it) is run twice per case, once as is and once with `injectCfScaleSegmentProbe` returning a unique sentinel ARGB
 * per segment evaluation, so each painted cell maps to the exact segment (value, thresholds, colours, cfvo kinds)
 * production evaluated for it. The baseline candidate (production's truncation) must reproduce production at every
 * included cell and Excel at every one but T4 and V3, else the harness is wrong (checked, and pinned by
 * tests/xlsx-cf-scale-num-segment-candidates.test.ts).
 *
 * Run: `npx tsx scripts/cf-refs/num-segment-candidates.ts` (prints the tables of docs/validation/xlsx-cf-recipes-stage2.md §8).
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { injectCfScaleSegmentProbe, interpolateScaleColor, numSegmentUnmeasured, type CfNumSegmentReading, type CfScaleSegment } from '../../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../../src/xlsx/cf/model'
import { caseDocumentEvaluationOf, caseParseOptionsOf } from './case-clock'
import type { CfManifest, CfOracleCase } from './cases'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

/** One captured colour-scale cell and the segment production evaluated for it. */
export interface SegmentCell {
  caseId: string
  ref: string
  segment: CfScaleSegment
  /** Excel's DisplayFormat fill (`FFRRGGBB`). */
  excel: string
  /** `confirmed`: production paints Excel's fill; `withheld`: production withholds it as the num-segment shape. */
  status: 'confirmed' | 'withheld'
}

export interface SegmentCollection {
  cells: SegmentCell[]
  /** Captured cells under a colour scale that are not in `cells`, with why (another gate withholds them, no fill observation). */
  excluded: Array<{ caseId: string; ref: string; reason: string }>
  /** Harness faults: an included cell whose production fill is neither Excel's nor the num-segment withhold (must be empty). */
  faults: string[]
}

const hex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0')
const rc = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}

/** Excel's DisplayFormat fill per ref: Stage 1/2 `observedCells`, Stage 5 `formulaObservations` (hidden cells skipped). */
function excelFills(c: CfOracleCase, rootDir: string): Map<string, string> | undefined {
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap || !existsSync(resolve(rootDir, cap.sidecarPath))) return undefined
  const sc = JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as {
    observedCells?: Array<{ ref: string; fillSource?: string; fillRgb?: string; hidden?: boolean }>
    formulaObservations?: Record<string, { displayFormatFill?: string }>
  }
  const out = new Map<string, string>()
  for (const o of sc.observedCells ?? []) if (o.fillSource === 'excel-display-format' && !o.hidden && typeof o.fillRgb === 'string') out.set(o.ref.toUpperCase(), o.fillRgb.toUpperCase())
  for (const [ref, o] of Object.entries(sc.formulaObservations ?? {})) if (typeof o.displayFormatFill === 'string' && /^[A-Z]+\d+$/.test(ref)) out.set(ref, o.displayFormatFill.toUpperCase())
  return out.size > 0 ? out : undefined
}

const pointOf = (kind: CfScaleSegment['loKind'], value: number): CfScalePoint =>
  kind === 'min' || kind === 'max' ? { bound: { type: kind }, argb: 0 } : { bound: { type: kind.split(':')[0] as 'num', value, ...(kind.endsWith(':formula') ? { resolved: 'formula' as const } : {}) }, argb: 0 }
/** Whether production withholds this segment as the num-segment shape (`numSegmentUnmeasured`). */
export const withheldShape = (s: CfScaleSegment): boolean =>
  s.colours === 3 && numSegmentUnmeasured(pointOf(s.loKind, s.lo), pointOf(s.hiKind, s.hi), s.value, s.lo, s.hi, s.fraction)

/** Every captured colour-scale cell of every measured case, mapped to its segment through production's own pipeline. */
export async function collectSegmentCells(manifest: CfManifest, rootDir: string = REPO_ROOT): Promise<SegmentCollection> {
  const out: SegmentCollection = { cells: [], excluded: [], faults: [] }
  for (const c of manifest.cases) {
    if (c.status !== 'measured' || !c.markers.some(m => m.ruleType === 'colorScale')) continue
    const fills = excelFills(c, rootDir)
    if (!fills) { out.excluded.push({ caseId: c.id, ref: '*', reason: 'no DisplayFormat fill observation in the capture (bar/icon sidecar)' }); continue }
    const bytes = new Uint8Array(readFileSync(resolve(rootDir, c.saved)))
    const doc = await parseXlsx(await OfficePackage.load(bytes), c.stage === 5 ? caseParseOptionsOf(c) : {})
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const options = { evidence: cfRuntimeEvidence, evaluation: c.stage === 5 ? caseDocumentEvaluationOf(c, doc) : cfDocumentEvaluationOf(doc), documentSheets: doc.sheets }
    const paintAt = async (): Promise<Map<string, string | undefined>> => {
      const paint = await prepareCfPaint(sheet, options)
      const m = new Map<string, string | undefined>()
      for (const ref of fills.keys()) { const { row, col } = rc(ref); m.set(ref, paint?.layer.styleAt(row, col, undefined)?.fillColor?.toUpperCase()) }
      paint?.dispose()
      return m
    }
    const production = await paintAt()
    const issued = new Map<string, CfScaleSegment>()
    let n = 0
    injectCfScaleSegmentProbe(s => {
      const argb = (0xff000000 | (0x5a0000 + ++n)) >>> 0
      issued.set(hex(argb), s)
      return { kind: 'ready', value: argb }
    })
    let probed: Map<string, string | undefined>
    try { probed = await paintAt() } finally { injectCfScaleSegmentProbe(null) }
    for (const [ref, excel] of fills) {
      const seg = probed.get(ref) !== undefined ? issued.get(probed.get(ref)!) : undefined
      const scaleCovered = c.markers.some(m => m.ruleType === 'colorScale' && sqrefCovers(m.sqref, ref))
      if (!seg) { if (scaleCovered) out.excluded.push({ caseId: c.id, ref, reason: production.get(ref) === undefined ? 'withheld before interpolation (another gate)' : 'painted without a segment evaluation (equal thresholds return an endpoint colour, or another rule)' }); continue }
      const prod = production.get(ref)
      const truncate = hex(interpolateScaleColor(seg.c0, seg.c1, seg.fraction))
      if (prod === undefined && withheldShape(seg)) { out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'withheld' }); continue }
      if (prod === undefined) { out.excluded.push({ caseId: c.id, ref, reason: 'interpolated, then withheld by a later gate (composition)' }); continue }
      if (prod !== truncate) { out.excluded.push({ caseId: c.id, ref, reason: `painted ${prod} by another rule over the scale` }); continue }
      if (prod !== excel) { out.faults.push(`${c.id}:${ref} production ${prod} is not Excel's ${excel}`); continue }
      out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'confirmed' })
    }
  }
  return out
}

function sqrefCovers(sqref: string, ref: string): boolean {
  const { row, col } = rc(ref)
  return sqref.split(/\s+/).some(part => {
    const [a, b = a] = part.split(':')
    const p = rc(a!.replace(/\$/g, '')), q = rc(b!.replace(/\$/g, ''))
    return row >= Math.min(p.row, q.row) && row <= Math.max(p.row, q.row) && col >= Math.min(p.col, q.col) && col <= Math.max(p.col, q.col)
  })
}

// ---------------------------------------------------------------------------------------------------------------------
// Candidates: a fraction formulation crossed with a channel formulation (and the fixed-point family). Each is a closed,
// deterministic formula of the segment; every one clamps the fraction to [0, 1] as production does.
// ---------------------------------------------------------------------------------------------------------------------

const clamp01 = (f: number): number => Math.max(0, Math.min(1, f))
const f32 = Math.fround
/** The largest double below `x` (0 < x). */
function nextDown(x: number): number {
  const buf = new Float64Array([x]), bits = new BigInt64Array(buf.buffer)
  bits[0] = bits[0]! - 1n
  return buf[0]!
}
const gSpan = (s: CfScaleSegment) => { const tMin = s.thresholds[0]!, tMax = s.thresholds[s.thresholds.length - 1]!; return (x: number): number => (x - tMin) / (tMax - tMin) }

/** Fraction formulations: the segment's position of the value. */
export const FRACTIONS: Readonly<Record<string, (s: CfScaleSegment) => number>> = {
  exact: s => (s.value - s.lo) / (s.hi - s.lo),
  split: s => s.value / (s.hi - s.lo) - s.lo / (s.hi - s.lo),
  reciprocal: s => (s.value - s.lo) * (1 / (s.hi - s.lo)),
  complement: s => 1 - (s.hi - s.value) / (s.hi - s.lo),
  'full-span': s => { const g = gSpan(s); return (g(s.value) - g(s.lo)) / (g(s.hi) - g(s.lo)) },
  'full-span-reciprocal': s => { const g = gSpan(s); return (g(s.value) - g(s.lo)) * (1 / (g(s.hi) - g(s.lo))) },
  'full-span-midpoint': s => { const g = gSpan(s); const gv = g(s.value); return s.segment === 0 ? gv / g(s.hi) : (gv - g(s.lo)) / (1 - g(s.lo)) },
  float32: s => f32((s.value - s.lo) / (s.hi - s.lo)),
  'float32-ops': s => f32(f32(f32(s.value) - f32(s.lo)) / f32(f32(s.hi) - f32(s.lo))),
  'percent-roundtrip': s => (((s.value - s.lo) / (s.hi - s.lo)) * 100) / 100,
  'percent-first': s => ((s.value - s.lo) * 100) / (s.hi - s.lo) / 100,
  'percent-floor': s => Math.floor(((s.value - s.lo) / (s.hi - s.lo)) * 100) / 100,
  'percent-round': s => Math.round(((s.value - s.lo) / (s.hi - s.lo)) * 100) / 100,
  'next-down': s => { const f = (s.value - s.lo) / (s.hi - s.lo); return f > 0 && f < 1 ? nextDown(f) : f },
  ...Object.fromEntries([100, 128, 255, 256, 512, 1000, 1024, 4096, 10000, 65536].flatMap(N => [
    [`fixed-floor-${N}`, (s: CfScaleSegment) => Math.floor(((s.value - s.lo) / (s.hi - s.lo)) * N) / N],
    [`fixed-round-${N}`, (s: CfScaleSegment) => Math.round(((s.value - s.lo) / (s.hi - s.lo)) * N) / N],
  ])),
}

const GUARD = 1e-11
const sgn = (d: number): number => (d >= 0 ? 1 : -1)
/** Rounds `x` toward `toward`: floor when `toward` is below, ceil when above. */
const towardOf = (x: number, toward: number): number => (toward <= x ? Math.floor(x) : Math.ceil(x))

/** Channel formulations: one 0..255 channel from its endpoints and the fraction (and the segment, for the ratio form). */
export const CHANNELS: Readonly<Record<string, (c0: number, c1: number, f: number, s: CfScaleSegment) => number>> = {
  // production: the step magnitude truncated, toward the segment's start colour, with the 1e-11 guard
  'step-truncate': (c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(f * Math.abs(c1 - c0) + GUARD),
  'step-truncate-unguarded': (c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(f * Math.abs(c1 - c0)),
  'value-floor': (c0, c1, f) => Math.floor(c0 + f * (c1 - c0)),
  'value-floor-guarded': (c0, c1, f) => Math.floor(c0 + f * (c1 - c0) + GUARD),
  'value-ceil': (c0, c1, f) => Math.ceil(c0 + f * (c1 - c0)),
  'value-round': (c0, c1, f) => Math.round(c0 + f * (c1 - c0)),
  'value-toward-start': (c0, c1, f) => towardOf(c0 + f * (c1 - c0), c0),
  'value-toward-end': (c0, c1, f) => towardOf(c0 + f * (c1 - c0), c1),
  'convex-floor': (c0, c1, f) => Math.floor(c0 * (1 - f) + c1 * f),
  'convex-toward-start': (c0, c1, f) => towardOf(c0 * (1 - f) + c1 * f, c0),
  'convex-round': (c0, c1, f) => Math.round(c0 * (1 - f) + c1 * f),
  'other-end-toward-start': (c0, c1, f) => towardOf(c1 + (1 - f) * (c0 - c1), c0),
  'other-end-toward-end': (c0, c1, f) => towardOf(c1 + (1 - f) * (c0 - c1), c1),
  'other-end-floor': (c0, c1, f) => Math.floor(c1 + (1 - f) * (c0 - c1)),
  'other-end-round': (c0, c1, f) => Math.round(c1 + (1 - f) * (c0 - c1)),
  'other-end-step-truncate': (c0, c1, f) => c1 - sgn(c1 - c0) * Math.ceil((1 - f) * Math.abs(c1 - c0) - GUARD),
  'step-toward-end': (c0, c1, f) => c1 - sgn(c1 - c0) * Math.floor((1 - f) * Math.abs(c1 - c0) + GUARD),
  'ratio-step-truncate': (c0, c1, _f, s) => c0 + sgn(c1 - c0) * Math.floor((Math.abs(c1 - c0) * (s.value - s.lo)) / (s.hi - s.lo) + GUARD),
  'ratio-step-truncate-unguarded': (c0, c1, _f, s) => c0 + sgn(c1 - c0) * Math.floor((Math.abs(c1 - c0) * (s.value - s.lo)) / (s.hi - s.lo)),
  'float32-step-truncate': (c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(f32(f32(f) * Math.abs(c1 - c0))),
  'float32-value-toward-start': (c0, c1, f) => towardOf(f32(c0 + f32(f32(f) * (c1 - c0))), c0),
  'float32-value-floor': (c0, c1, f) => Math.floor(f32(c0 + f32(f32(f) * (c1 - c0)))),
}

/** The fixed-point family: k = floor(f·N) (exact fraction), then integer arithmetic. */
const FIXED_N = [64, 100, 128, 255, 256, 512, 1000, 1024, 4096, 65536]
const FIXED: Readonly<Record<string, (c0: number, c1: number, f: number) => number>> = Object.fromEntries(FIXED_N.flatMap(N => [
  [`int${N}-step-truncate`, (c0: number, c1: number, f: number) => c0 + sgn(c1 - c0) * Math.floor((Math.abs(c1 - c0) * Math.floor(f * N)) / N)],
  [`int${N}-convex-floor`, (c0: number, c1: number, f: number) => { const k = Math.floor(f * N); return Math.floor((c0 * (N - k) + c1 * k) / N) }],
  [`int${N}-convex-round`, (c0: number, c1: number, f: number) => { const k = Math.floor(f * N); return Math.round((c0 * (N - k) + c1 * k) / N) }],
  [`int${N}-round-k-step-truncate`, (c0: number, c1: number, f: number) => c0 + sgn(c1 - c0) * Math.floor((Math.abs(c1 - c0) * Math.round(f * N)) / N)],
]))

/**
 * The sweeps (advisor-requested breadth): the fixed-point grid at every N from 2 to 5000 and at 2^13..2^40, and the
 * scaled round trip `(f·K)/K` in three operation orders. Each member is a closed formula; the step is truncated, unguarded.
 */
export const SWEEP_N: readonly number[] = [...Array.from({ length: 4999 }, (_, i) => i + 2), ...Array.from({ length: 28 }, (_, i) => 2 ** (i + 13))]
export const SWEEP_K: readonly number[] = [2, 3, 7, 10, 12, 16, 24, 60, 100, 255, 256, 360, 1000, 1024, 10000, 65535, 65536, 1e6]
const SWEEPS: Readonly<Record<string, (f: number, d: number) => number>> = {
  ...Object.fromEntries(SWEEP_N.map(N => [`sweep-fixed-${N}`, (f: number, d: number) => Math.floor((Math.floor(f * N) / N) * d)])),
  ...Object.fromEntries(SWEEP_K.flatMap(K => [
    [`sweep-roundtrip-${K}`, (f: number, d: number) => Math.floor(((f * K) / K) * d)],
    [`sweep-roundtrip-reciprocal-${K}`, (f: number, d: number) => Math.floor(f * K * (1 / K) * d)],
    [`sweep-roundtrip-after-${K}`, (f: number, d: number) => Math.floor((f * K * d) / K)],
  ])),
}

export interface Candidate { id: string; colour: (s: CfScaleSegment) => number }

const channelsOf = (argb: number): number[] => [(argb >>> 24) & 0xff, (argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]
function composeArgb(s: CfScaleSegment, channel: (c0: number, c1: number) => number): number {
  const a = channelsOf(s.c0), b = channelsOf(s.c1)
  // Alpha is never observed (DisplayFormat fills are opaque); every captured endpoint is FF.
  const out = [a[0]!, ...[1, 2, 3].map(i => channel(a[i]!, b[i]!))]
  return ((out[0]! << 24) | (out[1]! << 16) | (out[2]! << 8) | out[3]!) >>> 0
}

/** Every candidate: fraction × channel, plus the fixed-point family over the exact fraction. */
export const CANDIDATES: readonly Candidate[] = [
  ...Object.entries(FRACTIONS).flatMap(([fid, frac]) => Object.entries(CHANNELS)
    // The ratio channels read the segment directly (no fraction): only with the exact fraction, once.
    .filter(([cid]) => !cid.startsWith('ratio-') || fid === 'exact')
    .map(([cid, ch]) => ({ id: cid.startsWith('ratio-') ? cid : `${fid}/${cid}`, colour: (s: CfScaleSegment) => { const f = clamp01(frac(s)); return composeArgb(s, (c0, c1) => ch(c0, c1, f, s)) } }))),
  ...Object.entries(FIXED).map(([id, ch]) => ({ id, colour: (s: CfScaleSegment) => { const f = clamp01(FRACTIONS.exact!(s)); return composeArgb(s, (c0, c1) => ch(c0, c1, f)) } })),
  ...Object.entries(SWEEPS).map(([id, step]) => ({ id, colour: (s: CfScaleSegment) => { const f = clamp01(FRACTIONS.exact!(s)); return composeArgb(s, (c0, c1) => c0 + sgn(c1 - c0) * step(f, Math.abs(c1 - c0))) } })),
]
export const BASELINE_ID = 'exact/step-truncate'
/** The candidate each production reading (`CF_NUM_SEGMENT_READINGS`, colorScale.ts) implements: its class's representative. */
export const NUM_SEGMENT_READING_CANDIDATE: Readonly<Record<CfNumSegmentReading, string>> = {
  truncate: BASELINE_ID,
  'percent-roundtrip': 'percent-roundtrip/step-truncate-unguarded',
  'fixed-256': 'sweep-fixed-256',
  'fixed-1000': 'sweep-fixed-1000',
  'fixed-308': 'sweep-fixed-308',
  'fixed-700': 'sweep-fixed-700',
}

/** Where a candidate replaces production's arithmetic; outside it, production's (already Excel-equal) colour stands. */
export const SCOPES: Readonly<Record<string, (s: CfScaleSegment) => boolean>> = {
  all: () => true,
  'literal-num': s => s.loKind === 'num' && s.hiKind === 'num',
  '3-colour': s => s.colours === 3,
  '3-colour-literal-num': s => s.colours === 3 && s.loKind === 'num' && s.hiKind === 'num',
}

export interface CandidateResult { id: string; scope: string; inScope: number; mismatches: string[]; t4: boolean; v3: boolean }
const T4 = 'stage2-scale-scope-disjoint:T4', V3 = 'stage2-scale-scope-disjoint:V3'

export function evaluateCandidates(cells: readonly SegmentCell[], candidates: readonly Candidate[] = CANDIDATES): CandidateResult[] {
  const out: CandidateResult[] = []
  for (const cand of candidates) for (const [scope, inScope] of Object.entries(SCOPES)) {
    const scoped = cells.filter(c => inScope(c.segment))
    const mismatches: string[] = []
    let t4 = false, v3 = false
    for (const c of scoped) {
      const got = hex(cand.colour(c.segment))
      const key = `${c.caseId}:${c.ref}`
      if (got === c.excel) { if (key === T4) t4 = true; if (key === V3) v3 = true } else mismatches.push(`${key} ${got}≠${c.excel}`)
    }
    out.push({ id: cand.id, scope, inScope: scoped.length, mismatches, t4, v3 })
  }
  return out
}

/** The effective fraction window each cell's observed fill allows under step truncation toward the start colour. */
export function fractionWindow(c: SegmentCell): { lo: number; hi: number } {
  const a = channelsOf(c.segment.c0), b = channelsOf(c.segment.c1), o = channelsOf(Number.parseInt(c.excel, 16))
  let lo = 0, hi = Number.POSITIVE_INFINITY
  for (const i of [1, 2, 3]) {
    const d = Math.abs(b[i]! - a[i]!)
    if (d === 0) continue
    const step = Math.abs(o[i]! - a[i]!)
    lo = Math.max(lo, step / d)
    hi = Math.min(hi, step === d ? Number.POSITIVE_INFINITY : (step + 1) / d)
  }
  return { lo, hi }
}

async function main(): Promise<void> {
  const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const { cells, excluded, faults } = await collectSegmentCells(manifest)
  const byCase = new Map<string, { confirmed: number; withheld: number }>()
  for (const c of cells) { const e = byCase.get(c.caseId) ?? { confirmed: 0, withheld: 0 }; e[c.status]++; byCase.set(c.caseId, e) }
  console.log(`## Cells\n\n${cells.length} captured colour-scale cells (${cells.filter(c => c.status === 'confirmed').length} confirmed, ${cells.filter(c => c.status === 'withheld').length} withheld as the num-segment shape); ${excluded.length} excluded; ${faults.length} harness faults.\n`)
  console.log('| Case | Confirmed | Withheld (num-segment) |\n|---|---|---|')
  for (const [id, e] of byCase) console.log(`| \`${id}\` | ${e.confirmed} | ${e.withheld} |`)
  const exReasons = new Map<string, number>()
  for (const e of excluded) exReasons.set(`${e.caseId}: ${e.reason}`, (exReasons.get(`${e.caseId}: ${e.reason}`) ?? 0) + 1)
  console.log('\nExcluded:')
  for (const [r, n] of exReasons) console.log(`- ${r} (${n})`)
  for (const f of faults) console.log(`FAULT ${f}`)
  const interior = (c: SegmentCell): boolean => c.segment.value > c.segment.lo && c.segment.value < c.segment.hi
  const classOf = (c: SegmentCell): string => `${c.segment.colours}-colour ${c.segment.loKind}/${c.segment.hiKind}`
  const classes = new Map<string, { cells: number; interior: number; whole: number }>()
  for (const c of cells) {
    const e = classes.get(classOf(c)) ?? { cells: 0, interior: 0, whole: 0 }
    e.cells++
    if (interior(c)) {
      e.interior++
      const w = fractionWindow(c)
      if (w.lo === FRACTIONS.exact!(c.segment)) e.whole++
    }
    classes.set(classOf(c), e)
  }
  console.log('\n| Segment class | Cells | Strictly inside | Whole-product cells (constrain a lowered fraction) |\n|---|---|---|---|')
  for (const [k, e] of [...classes].sort()) console.log(`| ${k} | ${e.cells} | ${e.interior} | ${e.whole} |`)
  console.log('\nWithheld cells (the effective-fraction window Excel\'s fill allows, under step truncation):\n')
  console.log('| Cell | Segment | Value | f | Window | Excel | Production |\n|---|---|---|---|---|---|---|')
  for (const c of cells.filter(x => x.status === 'withheld')) {
    const w = fractionWindow(c)
    console.log(`| ${c.caseId.replace('stage2-scale-', '')}:${c.ref} | ${c.segment.lo}..${c.segment.hi} (${c.segment.colours}-colour, seg ${c.segment.segment}) | ${c.segment.value} | ${FRACTIONS.exact!(c.segment)} | [${w.lo.toFixed(6)}, ${w.hi === Infinity ? '∞' : w.hi.toFixed(6)}) | ${c.excel} | ${hex(interpolateScaleColor(c.segment.c0, c.segment.c1, c.segment.fraction))} |`)
  }
  const results = evaluateCandidates(cells)
  const base = results.find(r => r.id === BASELINE_ID && r.scope === 'all')!
  console.log(`\nBaseline ${BASELINE_ID} (production): ${base.mismatches.length} mismatches: ${base.mismatches.join('; ')}`)
  console.log(`\n## Candidates\n\n${CANDIDATES.length} candidates × ${Object.keys(SCOPES).length} scopes.\n`)
  const survivors = results.filter(r => r.mismatches.length === 0)
  console.log(`Survivors (every in-scope cell, T4 and V3 included): ${survivors.length} candidate × scope pairs\n`)
  console.log('| Scope | In-scope cells | Survivors | Grid (fraction/channel) and fixed-point survivors | Sweep survivors |\n|---|---|---|---|---|')
  for (const scope of Object.keys(SCOPES)) {
    const ids = survivors.filter(r => r.scope === scope).map(r => r.id)
    const n = results.find(r => r.scope === scope)!.inScope
    const grid = ids.filter(id => !id.startsWith('sweep-'))
    const fixedN = ids.filter(id => /^sweep-fixed-\d+$/.test(id)).map(id => Number(id.slice('sweep-fixed-'.length)))
    const rt = ids.filter(id => id.startsWith('sweep-roundtrip'))
    const fixedText = fixedN.length === 0 ? '' : `fixed-point N: ${fixedN.length} of ${SWEEP_N.length} (smallest ${fixedN.slice(0, 6).join(', ')}, …; every 2^e from ${fixedN.filter(x => x > 5000).length > 0 ? `2^${Math.log2(Math.min(...fixedN.filter(x => x > 5000)))}` : '—'})`
    const byFraction = new Map<string, string[]>()
    for (const id of grid) { const [f, ch] = id.includes('/') ? id.split('/') as [string, string] : [id, '']; byFraction.set(f, [...(byFraction.get(f) ?? []), ch]) }
    const gridText = [...byFraction].map(([f, chs]) => (chs.length === 1 && chs[0] === '' ? f : `${f}/{${chs.join(', ')}}`)).join('; ')
    console.log(`| ${scope} | ${n} | ${ids.length} | ${gridText || '—'} | ${[fixedText, rt.join(', ')].filter(Boolean).join('; ') || '—'} |`)
  }
  console.log('\nGrid candidates reproducing both T4 and V3, with mismatches elsewhere (fewest first, 40 shown):\n')
  console.log('| Candidate | Scope | In scope | Mismatches | First |\n|---|---|---|---|---|')
  for (const r of results.filter(x => x.t4 && x.v3 && x.mismatches.length > 0 && !x.id.startsWith('sweep-')).sort((a, b) => a.mismatches.length - b.mismatches.length).slice(0, 40)) console.log(`| ${r.id} | ${r.scope} | ${r.inScope} | ${r.mismatches.length} | ${r.mismatches[0]} |`)
  const one = results.filter(x => x.t4 !== x.v3)
  console.log(`\nCandidate × scope pairs reproducing exactly one of T4/V3: ${one.length}; neither: ${results.filter(x => !x.t4 && !x.v3).length}; both: ${results.filter(x => x.t4 && x.v3).length}.`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
