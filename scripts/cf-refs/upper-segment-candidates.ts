/**
 * The upper-segment follow-up (brief cf-upper-segment, Phase 1a): the census of every captured interior cell of a 3-colour
 * UPPER segment, and a candidate search for the arithmetic Excel 16.115 uses there.
 *
 * Cells: `collectSegmentCells` (scripts/cf-refs/num-segment-candidates.ts, `{ include: 'all' }`): every measured case with a
 * colour scale, Stage 2 and the Stage 5 painted colour-scale cells, each mapped to the segment production evaluated for it
 * (after the midpoint-above-max re-sort). Every capture is included, stage2-scale-literal-lower (a3f85309) and its 16 upper
 * probes among them. A cell is in the census when it is strictly inside the upper segment of a 3-colour scale (lo < v < hi).
 *
 * Per channel with d = c1 − c0 != 0: p = f·|d| (production's `safeFraction` double times the integer |d|, one IEEE
 * multiplication), its class (`integer`, `below` / `above`: within 1e-11 of a whole number k without being one, `clear`),
 * guarded truncation's step floor(p + 1e-11) and Excel's step (Excel's channel moved from c0, signed in the step direction);
 * the offset is Excel's step minus truncation's. A cell's class: `truncation` (Excel = guarded truncation), `raw-floor`
 * (Excel = floor(p) on every channel, differing from truncation: only `below` channels moved), `minus-one` (every channel
 * moved is one step below truncation and some `integer` or `above` channel moved), `other`.
 *
 * Search: a candidate is a closed deterministic formula of the segment. It replaces production's arithmetic strictly inside
 * the upper segment of a 3-colour scale only, so every lower-segment and 2-colour cell keeps production's painting by
 * construction (lower behaviour unchanged). A candidate SURVIVES when it reproduces Excel's DisplayFormat fill at every census
 * cell (all channels); undefined (fail closed) is a mismatch. Pools:
 * - `num-segment`: the 5803 candidates of the first search (`numSegmentCandidates`, recipe §8.1);
 * - `upper-window`: the 48,603 of the extended and window searches (`upperWindowCandidates`: origin, f32, precision, rational,
 *   form, threshold and the two window readings; recipe §9.1, §10.1);
 * - `literal-lower`: `step-plus-one` (`literalLowerStepColour`, recipe §12.2);
 * - `position` (new, this brief): Excel's position on the whole scale as a double, the segment index plus the fraction (or
 *   a rescaling of it: [0, 2], [0, 1] with the midpoint at 0.5, percent), the local fraction recovered from it by the
 *   inverse operations, then the step. The position may be built from production's f, from the complement from the maximum
 *   r = (hi − v)/(hi − lo), or with the segment index added to the offset before the division. Steps: raw floor, guarded
 *   floor, round, strictly below, the position product floor(t·|d|) − i·|d|, and value lerps rounded toward the start colour.
 *
 * Nothing here is registered or reachable from production: survivors are hypotheses to pre-register for a NEW capture.
 *
 * Run: `npx tsx scripts/cf-refs/upper-segment-candidates.ts` (about a minute; the search as run before the capture of
 * stage2-scale-upper-segment, recipe §13.1); `--all` includes that capture (recipe §13.7). Pinned by
 * tests/xlsx-cf-scale-upper-segment-candidates.test.ts.
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { CF_TRUNCATION_GUARD, interpolateScaleColor, literalLowerStepColour, type CfScaleSegment } from '../../src/xlsx/cf/colorScale'
import type { CfManifest } from './cases'
import { collectSegmentCells, type SegmentCell, type SegmentCollection } from './num-segment-candidates'
import { numSegmentCandidates } from './num-segment-formulas'
import { upperWindowCandidates } from './upper-window-candidates'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const f64 = new Float64Array(1), u64 = new BigUint64Array(f64.buffer)
/** The IEEE bits of a double, 16 hex digits. */
export const bitsOf = (x: number): string => { f64[0] = x; return u64[0]!.toString(16).padStart(16, '0') }
const sgn = (d: number): number => (d >= 0 ? 1 : -1)

// ---------------------------------------------------------------------------------------------------------------------
// Census.
// ---------------------------------------------------------------------------------------------------------------------

export type ProductClass = 'integer' | 'below' | 'above' | 'clear'
export function productClass(p: number): ProductClass {
  if (Number.isInteger(p)) return 'integer'
  const k = Math.round(p)
  if (Math.abs(p - k) <= CF_TRUNCATION_GUARD) return p < k ? 'below' : 'above'
  return 'clear'
}

export interface CensusChannel {
  channel: 'R' | 'G' | 'B'
  c0: number
  c1: number
  d: number
  /** p = f·|d| as production computes it. */
  p: number
  pHex: string
  cls: ProductClass | 'zero'
  /** Guarded truncation's step floor(p + 1e-11) and raw floor's floor(p) (0 when d = 0). */
  truncStep: number
  rawStep: number
  /** Excel's channel and its step from c0 in the step direction. */
  excel: number
  excelStep: number
  /** excelStep − truncStep. */
  offset: number
}

export type CellClass = 'truncation' | 'raw-floor' | 'minus-one' | 'other'
export interface CensusRow {
  caseId: string
  ref: string
  value: number
  lo: number
  hi: number
  thresholds: readonly number[]
  f: number
  fHex: string
  bounds: string
  /** `c0>c1` (the segment's start colour, the midpoint's, then the maximum's). */
  palette: string
  /** Directions of the channels with d != 0. */
  order: 'asc' | 'desc' | 'both'
  /** `p-line`: every channel with d != 0 has the same |d|; `distinct`. */
  kind: 'p-line' | 'distinct'
  excel: string
  truncation: string
  status: SegmentCell['status']
  channels: CensusChannel[]
  cls: CellClass
  /** The segment production evaluated (candidates are scored on it). */
  segment: CfScaleSegment
}

const upperInterior = (s: CfScaleSegment): boolean => s.colours === 3 && s.segment === 1 && s.value > s.lo && s.value < s.hi

export function upperSegmentCensus(collection: SegmentCollection): CensusRow[] {
  const rows: CensusRow[] = []
  for (const x of collection.cells) {
    const s = x.segment
    if (!upperInterior(s)) continue
    const ex = Number.parseInt(x.excel, 16) >>> 0
    const channels: CensusChannel[] = ([['R', 16], ['G', 8], ['B', 0]] as const).map(([channel, shift]) => {
      const c0 = (s.c0 >>> shift) & 0xff, c1 = (s.c1 >>> shift) & 0xff, d = c1 - c0, ad = Math.abs(d)
      const p = s.fraction * ad, e = (ex >>> shift) & 0xff
      const truncStep = ad === 0 ? 0 : Math.floor(p + CF_TRUNCATION_GUARD), rawStep = ad === 0 ? 0 : Math.floor(p)
      const excelStep = ad === 0 ? e - c0 : (e - c0) * sgn(d)
      return { channel, c0, c1, d, p, pHex: bitsOf(p), cls: ad === 0 ? 'zero' : productClass(p), truncStep, rawStep, excel: e, excelStep, offset: excelStep - truncStep }
    })
    const live = channels.filter(c => c.d !== 0)
    const dirs = new Set(live.map(c => (c.d > 0 ? 'asc' : 'desc')))
    const truncation = hex(interpolateScaleColor(s.c0, s.c1, s.fraction))
    let cls: CellClass
    if (channels.every(c => c.offset === 0)) cls = 'truncation'
    else if (channels.every(c => c.excelStep === c.rawStep)) cls = 'raw-floor'
    else if (channels.every(c => c.offset === 0 || (c.offset === -1 && c.cls !== 'clear' && c.cls !== 'zero')) && channels.some(c => c.offset === -1 && (c.cls === 'integer' || c.cls === 'above'))) cls = 'minus-one'
    else cls = 'other'
    rows.push({
      caseId: x.caseId, ref: x.ref, value: s.value, lo: s.lo, hi: s.hi, thresholds: s.thresholds, f: s.fraction, fHex: bitsOf(s.fraction),
      bounds: `${s.loKind}/${s.hiKind}`, palette: `${hex(s.c0)}>${hex(s.c1)}`, order: dirs.size === 2 ? 'both' : dirs.has('desc') ? 'desc' : 'asc',
      kind: new Set(live.map(c => Math.abs(c.d))).size <= 1 ? 'p-line' : 'distinct', excel: x.excel, truncation, status: x.status, channels, cls, segment: s,
    })
  }
  return rows
}

/** Per-channel table: product class × Excel's offset from truncation (channels with d != 0). */
export function channelTable(rows: readonly CensusRow[]): Array<{ key: string; channels: number }> {
  const m = new Map<string, number>()
  for (const r of rows) for (const c of r.channels) if (c.d !== 0) { const k = `${c.cls} ${c.offset}`; m.set(k, (m.get(k) ?? 0) + 1) }
  return [...m].sort(([a], [b]) => (a < b ? -1 : 1)).map(([key, channels]) => ({ key, channels }))
}

export function countBy<T>(rows: readonly T[], key: (r: T) => string): Array<{ key: string; cells: number }> {
  const m = new Map<string, number>()
  for (const r of rows) m.set(key(r), (m.get(key(r)) ?? 0) + 1)
  return [...m].sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, cells]) => ({ key: k, cells }))
}

/**
 * Whether one decision per f double explains the census: every channel that is not `clear` moved (offset −1) or not
 * (offset 0) alike within a cell and across every cell with the same f double. Returns the f doubles with their outcome.
 */
export function decisionByDouble(rows: readonly CensusRow[]): { keys: Map<string, 'moved' | 'kept'>; mixedCells: string[]; inconsistent: string[] } {
  const keys = new Map<string, 'moved' | 'kept'>(), mixedCells: string[] = [], inconsistent = new Set<string>()
  for (const r of rows) {
    const near = r.channels.filter(c => c.cls === 'integer' || c.cls === 'below' || c.cls === 'above')
    if (near.length === 0) continue
    const outs = new Set(near.map(c => c.offset))
    if (outs.size !== 1 || ![0, -1].includes([...outs][0]!)) { mixedCells.push(`${r.caseId}:${r.ref}`); continue }
    const o = [...outs][0] === -1 ? 'moved' : 'kept'
    const prev = keys.get(r.fHex)
    if (prev !== undefined && prev !== o) inconsistent.add(r.fHex)
    keys.set(r.fHex, o)
  }
  return { keys, mixedCells, inconsistent: [...inconsistent] }
}

// ---------------------------------------------------------------------------------------------------------------------
// The new family: the position on the whole scale.
// ---------------------------------------------------------------------------------------------------------------------

/** How the position t is formed from the segment (segment index i = 1), and how the local fraction is recovered from it. */
type PositionForm = { t: (s: CfScaleSegment) => number; local: (t: number) => number; index: number; unit: number }
const fOf = (s: CfScaleSegment): number => s.fraction
const rOf = (s: CfScaleSegment): number => (s.hi - s.value) / (s.hi - s.lo)
export const POSITION_FORMS: Readonly<Record<string, PositionForm>> = {
  // [0, 2]: lower segment 0..1, upper 1..2.
  'index-plus-f': { t: s => 1 + fOf(s), local: t => t - 1, index: 1, unit: 1 },
  'f-plus-index': { t: s => fOf(s) + 1, local: t => t - 1, index: 1, unit: 1 },
  'index-plus-offset-quotient': { t: s => (s.value - s.lo + (s.hi - s.lo)) / (s.hi - s.lo), local: t => t - 1, index: 1, unit: 1 },
  'two-minus-r': { t: s => 2 - rOf(s), local: t => t - 1, index: 1, unit: 1 },
  'index-plus-one-minus-r': { t: s => 1 + (1 - rOf(s)), local: t => t - 1, index: 1, unit: 1 },
  // [0, 1] with the midpoint at 0.5.
  'half-plus-half-f': { t: s => 0.5 + 0.5 * fOf(s), local: t => (t - 0.5) * 2, index: 0.5, unit: 0.5 },
  'half-plus-half-f-div': { t: s => 0.5 + fOf(s) / 2, local: t => (t - 0.5) / 0.5, index: 0.5, unit: 0.5 },
  'index-plus-f-halved': { t: s => (1 + fOf(s)) / 2, local: t => 2 * t - 1, index: 0.5, unit: 0.5 },
  'one-minus-half-r': { t: s => 1 - 0.5 * rOf(s), local: t => (t - 0.5) * 2, index: 0.5, unit: 0.5 },
  // Percent of the whole scale, the midpoint at 50.
  'percent-50': { t: s => 50 + 50 * fOf(s), local: t => (t - 50) / 50, index: 50, unit: 50 },
  'percent-50-from-f100': { t: s => 50 + (fOf(s) * 100) / 2, local: t => (t - 50) / 50, index: 50, unit: 50 },
  // [0, 2] in percent (each segment 100).
  'percent-100': { t: s => 100 + 100 * fOf(s), local: t => (t - 100) / 100, index: 100, unit: 100 },
  // Other indexes: the threshold count (a segment index of 2), float32 position.
  'index-2-plus-f': { t: s => 2 + fOf(s), local: t => t - 2, index: 2, unit: 1 },
  'index-plus-f-f32': { t: s => Math.fround(1 + fOf(s)), local: t => Math.fround(t - 1), index: 1, unit: 1 },
  // The neighbourhood of the survivor: f from a reciprocal multiply or from the complement, another offset (0.5 + f), and
  // an explicit snap of f to the 2^-52 grid under each rounding rule (nearest-even is the survivor's own snap).
  'reciprocal-index-plus-f': { t: s => 1 + (s.value - s.lo) * (1 / (s.hi - s.lo)), local: t => t - 1, index: 1, unit: 1 },
  'complement-index-plus-f': { t: s => 1 + (1 - rOf(s)), local: t => t - 1, index: 1, unit: 1 },
  'half-plus-f': { t: s => 0.5 + fOf(s), local: t => t - 0.5, index: 0.5, unit: 1 },
  'snap-52-nearest-even': { t: s => snap52(fOf(s), 'even'), local: t => t, index: 0, unit: 1 },
  'snap-52-nearest-up': { t: s => snap52(fOf(s), 'up'), local: t => t, index: 0, unit: 1 },
  'snap-52-floor': { t: s => snap52(fOf(s), 'floor'), local: t => t, index: 0, unit: 1 },
  'snap-52-ceil': { t: s => snap52(fOf(s), 'ceil'), local: t => t, index: 0, unit: 1 },
}
/** f (0 <= f <= 1) on the 2^-52 grid: nearest with ties to even or up, floor or ceil (every step exact in doubles). */
export function snap52(f: number, mode: 'even' | 'up' | 'floor' | 'ceil'): number {
  const x = f * 2 ** 52, lo = Math.floor(x), frac = x - lo
  const n = mode === 'floor' ? lo : mode === 'ceil' ? Math.ceil(x) : frac > 0.5 ? lo + 1 : frac < 0.5 ? lo : mode === 'up' ? lo + 1 : lo % 2 === 0 ? lo : lo + 1
  return n / 2 ** 52
}
/** floor(l·d) in exact arithmetic (l a finite double >= 0, d a non-negative integer): no rounding of the product. */
export function exactFloorProduct(l: number, d: number): number {
  let n = l, e = 0n
  while (!Number.isInteger(n)) { n *= 2; e++ }
  return Number((BigInt(n) * BigInt(d)) >> e)
}
type PositionStep = (t: number, local: number, d: number, form: PositionForm) => number
export const POSITION_STEPS: Readonly<Record<string, PositionStep>> = {
  floor: (_t, l, d) => Math.floor(l * d),
  guarded: (_t, l, d) => Math.floor(l * d + CF_TRUNCATION_GUARD),
  round: (_t, l, d) => Math.round(l * d),
  'strictly-below': (_t, l, d) => Math.ceil(l * d) - 1,
  // The position product: floor(t·|d|/unit) − index·|d|/unit.
  'position-product': (t, _l, d, form) => Math.floor((t / form.unit) * d) - (form.index / form.unit) * d,
  'floor-f32-product': (_t, l, d) => Math.floor(Math.fround(l * d)),
  'floor-exact-product': (_t, l, d) => exactFloorProduct(l, d),
}
/** Value lerps from the recovered local fraction, rounded toward the start colour (c0 + l·D) or floored. */
type PositionValue = (c0: number, c1: number, l: number) => number
export const POSITION_VALUES: Readonly<Record<string, PositionValue>> = {
  'value-toward-start': (c0, c1, l) => { const x = c0 + l * (c1 - c0); return c1 >= c0 ? Math.floor(x) : Math.ceil(x) },
  'value-floor': (c0, c1, l) => Math.floor(c0 + l * (c1 - c0)),
  'convex-toward-start': (c0, c1, l) => { const x = c0 * (1 - l) + c1 * l; return c1 >= c0 ? Math.floor(x) : Math.ceil(x) },
}

export interface UpperSegmentCandidate {
  id: string
  pool: 'num-segment' | 'upper-window' | 'literal-lower' | 'position' | 'baseline'
  /** The candidate's colour for one upper-segment interior cell, or undefined (fail closed). */
  evaluate: (s: CfScaleSegment) => number | undefined
}

/** Per RGB channel from c0 (alpha kept: every captured endpoint is opaque); undefined unless every channel is 0..255. */
function perChannel(s: CfScaleSegment, ch: (c0: number, c1: number) => number | undefined): number | undefined {
  let out = (s.c0 >>> 24) & 0xff
  for (const shift of [16, 8, 0]) {
    const c0 = (s.c0 >>> shift) & 0xff, c1 = (s.c1 >>> shift) & 0xff
    const c = c0 === c1 ? c0 : ch(c0, c1)
    if (c === undefined || !Number.isInteger(c) || c < 0 || c > 255) return undefined
    out = (out << 8) | c
  }
  return out >>> 0
}

function positionCandidates(): UpperSegmentCandidate[] {
  const out: UpperSegmentCandidate[] = []
  for (const [fid, form] of Object.entries(POSITION_FORMS)) {
    const local = (s: CfScaleSegment): { t: number; l: number } | undefined => {
      if (![s.value, s.lo, s.hi].every(Number.isFinite) || !(s.hi > s.lo)) return undefined
      const t = form.t(s), l = Math.max(0, Math.min(1, form.local(t)))
      return Number.isFinite(t) && Number.isFinite(l) ? { t, l } : undefined
    }
    for (const [sid, step] of Object.entries(POSITION_STEPS)) out.push({
      id: `position/${fid}/${sid}`, pool: 'position',
      evaluate: s => {
        const x = local(s); if (!x) return undefined
        return perChannel(s, (c0, c1) => { const d = Math.abs(c1 - c0), k = step(x.t, x.l, d, form); return Number.isFinite(k) ? c0 + sgn(c1 - c0) * Math.max(0, Math.min(d, k)) : undefined })
      },
    })
    for (const [vid, value] of Object.entries(POSITION_VALUES)) out.push({
      id: `position/${fid}/${vid}`, pool: 'position',
      evaluate: s => { const x = local(s); return x ? perChannel(s, (c0, c1) => value(c0, c1, x.l)) : undefined },
    })
  }
  return out
}

let built: UpperSegmentCandidate[] | undefined
/** Every candidate: production's truncation (baseline), the earlier pools, step-plus-one and the position family. */
export function upperSegmentCandidates(): readonly UpperSegmentCandidate[] {
  return (built ??= [
    { id: 'baseline/truncate', pool: 'baseline', evaluate: s => interpolateScaleColor(s.c0, s.c1, s.fraction) },
    ...numSegmentCandidates().map((c): UpperSegmentCandidate => ({ id: c.id, pool: 'num-segment', evaluate: c.evaluate })),
    ...upperWindowCandidates().map((c): UpperSegmentCandidate => ({ id: c.id, pool: 'upper-window', evaluate: c.evaluate })),
    { id: 'literal-lower/step-plus-one', pool: 'literal-lower', evaluate: s => literalLowerStepColour(s.c0, s.c1, s.fraction) },
    ...positionCandidates(),
  ])
}

export interface CandidateResult { id: string; pool: UpperSegmentCandidate['pool']; mismatches: string[] }

/** Scores candidates over the census cells (key `caseId:ref`). */
export function evaluateCandidates(rows: readonly CensusRow[], candidates: readonly UpperSegmentCandidate[] = upperSegmentCandidates()): CandidateResult[] {
  return candidates.map(cand => ({
    id: cand.id, pool: cand.pool,
    mismatches: rows.filter(r => { const got = cand.evaluate(r.segment); return got === undefined || hex(got) !== r.excel }).map(r => `${r.caseId}:${r.ref}`),
  }))
}

// ---------------------------------------------------------------------------------------------------------------------
// Survivor classes: survivors grouped by their colour over a fixed probe grid (uncaptured inputs).
// ---------------------------------------------------------------------------------------------------------------------

/**
 * The probe grid: upper-segment segments with lo = 0, hi = q, v = p (integer operands: f the correctly rounded p/q) for
 * q = 2..64 and 0 < p < q, plus decimal-operand segments (lo = 0.1·a, span 0.1·b) and every double within 3 ulps of those
 * fractions (lo = 0, hi = 1, v the double), each at the palettes FF000000>FFFFFFFF and FFFFFFFF>FF000000 (|d| = 255),
 * FF000000>FF1E0F07 / FF1E0F07>FF000000 (|d| = 30, 15, 7), the P line FF14F028 / FFE61EFA (|d| = 210) and the default
 * palette's upper pair FFFFEB84 / FF63BE7B (|d| = 156, 45, 9), both orders. Deterministic.
 */
export function probeGrid(): CfScaleSegment[] {
  const out: CfScaleSegment[] = []
  const palettes: Array<[number, number]> = [[0xff000000, 0xffffffff], [0xffffffff, 0xff000000], [0xff000000, 0xff1e0f07], [0xff1e0f07, 0xff000000], [0xff14f028, 0xffe61efa], [0xffe61efa, 0xff14f028], [0xffffeb84, 0xff63be7b], [0xff63be7b, 0xffffeb84]]
  const push = (value: number, lo: number, hi: number): void => {
    const fraction = (value - lo) / (hi - lo)
    if (!(fraction > 0 && fraction < 1)) return
    for (const [c0, c1] of palettes) out.push({ value, lo, hi, fraction, c0, c1, colours: 3, segment: 1, loKind: 'num', hiKind: 'num', thresholds: [lo - 1, lo, hi] })
  }
  const step = (x: number, n: number): number => { f64[0] = x; u64[0] = u64[0]! + BigInt(n); return f64[0]! }
  for (let q = 2; q <= 64; q++) for (let p = 1; p < q; p++) {
    push(p, 0, q)
    push(0.1 * (q + p), 0.1 * q, 0.1 * q + 0.1 * q)
    for (let u = -3; u <= 3; u++) if (u !== 0) push(step(p / q, u), 0, 1)
  }
  return out
}

/** Survivors grouped by identical colours over `probeGrid()`; each class lists its member ids. */
export function survivorClasses(ids: readonly string[], candidates: readonly UpperSegmentCandidate[] = upperSegmentCandidates()): string[][] {
  const grid = probeGrid()
  const byKey = new Map<string, string[]>()
  for (const id of ids) {
    const c = candidates.find(x => x.id === id)!
    const key = grid.map(s => { const v = c.evaluate(s); return v === undefined ? 'u' : v.toString(36) }).join(',')
    byKey.set(key, [...(byKey.get(key) ?? []), id])
  }
  return [...byKey.values()]
}

// ---------------------------------------------------------------------------------------------------------------------
// Entry.
// ---------------------------------------------------------------------------------------------------------------------

/**
 * The capture that followed this search (stage2-scale-upper-segment, 2ee3735a): `{ include: 'pre-capture' }` (the default, the
 * search as run before the capture, recipe §13.1) leaves it out; `'all'` includes it (recipe §13.7).
 */
export const UPPER_SEGMENT_POST_CAPTURE_CASES: readonly string[] = ['stage2-scale-upper-segment']
/** Captures after this search's own follow-up, left out in every mode: the census is recorded as of 2ee3735a (recipe §13.7). */
export const UPPER_SEGMENT_LATER_CASES: readonly string[] = ['stage2-scale-formula-bounds', 'stage2-scale-percent-order']

export async function upperSegmentCensusOf(manifest?: CfManifest, rootDir: string = REPO_ROOT, opts: { include?: 'pre-capture' | 'all' } = {}): Promise<{ collection: SegmentCollection; rows: CensusRow[] }> {
  const m = manifest ?? (JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest)
  const all = await collectSegmentCells(m, rootDir, { include: 'all' })
  if (all.faults.length > 0) throw new Error(`upper-segment: harness faults ${all.faults.join('; ')}`)
  const keep = (x: { caseId: string }): boolean => !UPPER_SEGMENT_LATER_CASES.includes(x.caseId) && ((opts.include ?? 'pre-capture') === 'all' || !UPPER_SEGMENT_POST_CAPTURE_CASES.includes(x.caseId))
  const collection: SegmentCollection = { cells: all.cells.filter(keep), excluded: all.excluded.filter(keep), faults: all.faults }
  return { collection, rows: upperSegmentCensus(collection) }
}

async function main(): Promise<void> {
  const { collection, rows } = await upperSegmentCensusOf(undefined, REPO_ROOT, { include: process.argv.includes('--all') ? 'all' : 'pre-capture' })
  const short = (k: string): string => k.replace(/^stage\d-(scale-)?/, '')
  const groupB3 = collection.excluded.filter(e => e.group === 'b' && e.threeColour).length
  console.log(`## Census\n\n${collection.cells.length} captured colour-scale cells; ${collection.excluded.length} excluded (${collection.excluded.filter(e => e.group === 'b').length} by a production gate, ${groupB3} of them under a 3-colour scale); ${rows.length} upper-segment interior cells.\n`)
  const table = (title: string, key: (r: CensusRow) => string): void => { console.log(`### ${title}\n`); for (const t of countBy(rows, key)) console.log(`- ${t.key}: ${t.cells}`); console.log('') }
  table('By class', r => r.cls)
  table('By class × bounds', r => `${r.cls} | ${r.bounds}`)
  table('By class × colour order', r => `${r.cls} | ${r.order} ${r.palette}`)
  table('By class × palette kind', r => `${r.cls} | ${r.kind}`)
  table('By class × case', r => `${r.cls} | ${r.caseId}`)
  console.log('### Per channel (product class × offset)\n')
  for (const t of channelTable(rows)) console.log(`- ${t.key}: ${t.channels}`)
  const dec = decisionByDouble(rows)
  console.log(`\nDecision per f double: ${dec.keys.size} doubles (${[...dec.keys.values()].filter(v => v === 'moved').length} moved), ${dec.mixedCells.length} mixed cells, ${dec.inconsistent.length} inconsistent doubles.\n`)
  if (process.argv.includes('--rows')) {
    console.log('| Cell | Bounds | v | lo | hi | f (hex) | c0>c1 | per channel d: p (hex) class trunc→Excel | Excel | Class |\n|---|---|---|---|---|---|---|---|---|---|')
    for (const r of rows) console.log(`| ${short(r.caseId)}:${r.ref} | ${r.bounds} | ${r.value} | ${r.lo} | ${r.hi} | ${r.fHex} | ${r.palette} | ${r.channels.map(c => (c.d === 0 ? `${c.channel} 0` : `${c.channel} ${c.d}: ${c.pHex} ${c.cls} ${c.truncStep}→${c.excelStep}`)).join('; ')} | ${r.excel} | ${r.cls} |`)
  }
  const results = evaluateCandidates(rows)
  console.log('## Search\n\n| Pool | Candidates | Survivors | Fewest mismatches | Best (first 3) |\n|---|---|---|---|---|')
  for (const pool of ['baseline', 'num-segment', 'upper-window', 'literal-lower', 'position'] as const) {
    const rs = results.filter(r => r.pool === pool)
    const best = Math.min(...rs.map(r => r.mismatches.length))
    console.log(`| ${pool} | ${rs.length} | ${rs.filter(r => r.mismatches.length === 0).length} | ${best} | ${rs.filter(r => r.mismatches.length === best).slice(0, 3).map(r => `\`${r.id}\``).join(', ')} |`)
  }
  const survivors = results.filter(r => r.mismatches.length === 0).map(r => r.id)
  console.log(`\nTotal ${results.length} candidates; ${survivors.length} survivors.`)
  const classes = survivorClasses(survivors)
  console.log(`\n## Survivor classes over the probe grid (${probeGrid().length} segments): ${classes.length}\n`)
  classes.forEach((c, i) => console.log(`${i + 1}. ${c.map(x => `\`${x}\``).join(', ')}`))
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
