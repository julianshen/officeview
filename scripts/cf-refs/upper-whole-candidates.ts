/**
 * The upper-segment whole-product follow-up (brief cf-upperwhole, Phase 1 task 1): an extended candidate search, aimed at
 * hypotheses H1-H3 of docs/validation/xlsx-cf.md §14.z.15, over every captured colour-scale cell (stage2-scale-num-segment
 * included: its 19 moved cells and all its unmoved ones, and the painted colour-scale cells of the Stage 5 refs).
 *
 * Capture 377564e0 (stage2-scale-num-segment) drew Excel's truncation everywhere except 19 cells, all in the UPPER segment of
 * a 3-colour scale at a whole step product (f = 1/3, 2/3, 1/5, 1/7, 1/15), each one unit smaller per channel; T4 and V3 of
 * e448d831 are the same shape. The first search (scripts/cf-refs/num-segment-candidates.ts, 5803 candidates) found no
 * formula for the moved / unmoved split. This one adds families that search did not enumerate:
 * - `origin/…` (H1, H3): the upper-segment fraction from another origin or quantity (from the maximum backwards, from the
 *   scale minimum over the full span and remapped by the midpoint's position, in several operation orders, percent
 *   scaling, reciprocal multiplies) crossed with channel lerps written from either end (rounded toward either colour, in
 *   0..255, 0..1 and packed-integer channels);
 * - `f32/…` (H3): float32 rounding at every subset of the fraction's operations crossed with float32 rounding of the
 *   channel product;
 * - `precision/…` (H2): the fraction (or its complement from the maximum) rounded to a p-bit binary significand
 *   (nearest or truncated, p = 18..53), to n significant decimal digits (6..17) or to float32, and the step product
 *   rounded by a second such precision: the split as a function of the double representation, always a closed formula;
 * - `rational/…` (H3): exact-rational emulation (BigInt) of FMA-contracted lerps and of every operation at a 24-, 53- or
 *   64-bit significand (x87 extended), or exact.
 * Each candidate is a closed deterministic formula of the segment, evaluated at three scopes (the upper segment of a
 * 3-colour scale only; every 3-colour segment; every segment). Outside its scope production's truncation stands. A candidate
 * SURVIVES a scope when it reproduces Excel's DisplayFormat fill at every captured cell. Evaluation fails closed: any
 * non-finite intermediate (float32 conversions included) or a channel that is not a whole number in 0..255 is a mismatch.
 *
 * Nothing here is registered or reachable from production (brief hard rule): survivors would be hypotheses for a NEW capture.
 *
 * Run: `npx tsx scripts/cf-refs/upper-whole-candidates.ts` (every capture, stage2-scale-upper-whole included: recipe §9.5);
 * `--pre-capture` leaves that capture out and prints the tables of docs/validation/xlsx-cf-recipes-stage2.md §9.1.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { interpolateScaleColor, numSegmentUnmeasured, upperSegmentWholeProduct, type CfScaleSegment } from '../../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../../src/xlsx/cf/model'
import { s2UpperWholeCells, s2UpperWholeLayout, s2UpperWholeTruncation, S2_UPPER_WHOLE_ID, type CfManifest, type S2UpperWholeCell } from './cases'
import { collectSegmentCells, type SegmentCell } from './num-segment-candidates'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

// ---------------------------------------------------------------------------------------------------------------------
// Fail-closed arithmetic: every intermediate passes through `t.v`, which records a non-finite value.
// ---------------------------------------------------------------------------------------------------------------------

class Trap {
  bad = false
  v(x: number): number { if (!Number.isFinite(x)) this.bad = true; return x }
  /** A float32 conversion (finite doubles above float32's range become Infinity: trapped). */
  f(x: number): number { return this.v(Math.fround(this.v(x))) }
}
const sgn = (d: number): number => (d >= 0 ? 1 : -1)
const clamp01 = (f: number): number => Math.max(0, Math.min(1, f))
const channelsOf = (argb: number): [number, number, number] => [(argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]
const hex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0')
/** Rounds `x` toward `toward` (floor when it is below, ceil when above). */
const towardOf = (x: number, toward: number): number => (toward <= x ? Math.floor(x) : Math.ceil(x))

/** Packs three channels (alpha FF; every captured endpoint is opaque); undefined unless each is a whole number in 0..255. */
function pack(rgb: readonly number[]): number | undefined {
  if (!rgb.every(c => Number.isInteger(c) && c >= 0 && c <= 255)) return undefined
  return ((0xff << 24) | (rgb[0]! << 16) | (rgb[1]! << 8) | rgb[2]!) >>> 0
}

export interface UpperWholeCandidate {
  id: string
  family: 'origin' | 'f32' | 'precision' | 'rational'
  /** The candidate's colour for one segment, or undefined (fail closed: non-finite intermediate, invalid channel). */
  evaluate: (s: CfScaleSegment) => number | undefined
}

/** The raw differences every formulation may read, finite and with non-zero spans (else undefined before any arithmetic). */
function inputsOf(s: CfScaleSegment): { v: number; lo: number; hi: number; m: number } | undefined {
  const m = s.thresholds[0]!
  const xs = [s.value, s.lo, s.hi, m, s.value - s.lo, s.hi - s.lo, s.hi - s.value, s.value - m, s.lo - m, s.hi - m]
  if (!xs.every(Number.isFinite) || s.hi - s.lo === 0 || s.hi - m === 0) return undefined
  return { v: s.value, lo: s.lo, hi: s.hi, m }
}

// ---------------------------------------------------------------------------------------------------------------------
// Family `origin` (H1, H3): fraction from another origin x channel lerp from either end, double precision.
// ---------------------------------------------------------------------------------------------------------------------

type Fraction = (t: Trap, v: number, lo: number, hi: number, m: number) => number
export const ORIGIN_FRACTIONS: Readonly<Record<string, Fraction>> = {
  exact: (t, v, lo, hi) => t.v(t.v(v - lo) / t.v(hi - lo)),
  'from-max': (t, v, lo, hi) => t.v(1 - t.v(t.v(hi - v) / t.v(hi - lo))),
  'from-max-reciprocal': (t, v, lo, hi) => t.v(1 - t.v(t.v(hi - v) * t.v(1 / t.v(hi - lo)))),
  'from-max-percent': (t, v, lo, hi) => t.v(t.v(100 - t.v(t.v(t.v(hi - v) / t.v(hi - lo)) * 100)) / 100),
  'split': (t, v, lo, hi) => t.v(t.v(v / t.v(hi - lo)) - t.v(lo / t.v(hi - lo))),
  // From the scale minimum: the full-span position g, remapped with the midpoint's position gm.
  'g-remap': (t, v, lo, hi, m) => { const g = t.v(t.v(v - m) / t.v(hi - m)), gm = t.v(t.v(lo - m) / t.v(hi - m)); return t.v(t.v(g - gm) / t.v(1 - gm)) },
  'g-remap-reciprocal': (t, v, lo, hi, m) => { const g = t.v(t.v(v - m) / t.v(hi - m)), gm = t.v(t.v(lo - m) / t.v(hi - m)); return t.v(t.v(g - gm) * t.v(1 / t.v(1 - gm))) },
  'g-remap-from-max': (t, v, lo, hi, m) => { const g = t.v(t.v(v - m) / t.v(hi - m)), gm = t.v(t.v(lo - m) / t.v(hi - m)); return t.v(1 - t.v(t.v(1 - g) / t.v(1 - gm))) },
  'g-remap-top': (t, v, lo, hi, m) => { const g = t.v(t.v(v - m) / t.v(hi - m)), gm = t.v(t.v(lo - m) / t.v(hi - m)), gh = t.v(t.v(hi - m) / t.v(hi - m)); return t.v(t.v(g - gm) / t.v(gh - gm)) },
  'g-remap-percent': (t, v, lo, hi, m) => { const g = t.v(t.v(t.v(v - m) / t.v(hi - m)) * 100), gm = t.v(t.v(t.v(lo - m) / t.v(hi - m)) * 100); return t.v(t.v(g - gm) / t.v(100 - gm)) },
  'g-remap-scaled': (t, v, lo, hi, m) => { const k = t.v(1 / t.v(hi - m)); return t.v(t.v(t.v(t.v(v - m) * k) - t.v(t.v(lo - m) * k)) / t.v(1 - t.v(t.v(lo - m) * k))) },
  'g-offsets': (t, v, lo, hi, m) => t.v(t.v(t.v(v - m) - t.v(lo - m)) / t.v(t.v(hi - m) - t.v(lo - m))),
  // From the maximum backwards over the full span, remapped with the midpoint's distance from the maximum.
  'g-reverse': (t, v, lo, hi, m) => { const r = t.v(t.v(hi - v) / t.v(hi - m)), rm = t.v(t.v(hi - lo) / t.v(hi - m)); return t.v(1 - t.v(r / rm)) },
  'g-reverse-difference': (t, v, lo, hi, m) => { const r = t.v(t.v(hi - v) / t.v(hi - m)), rm = t.v(t.v(hi - lo) / t.v(hi - m)); return t.v(t.v(rm - r) / rm) },
  // The midpoint-relative position over the full span, rescaled to the segment.
  'mid-over-full': (t, v, lo, hi, m) => t.v(t.v(t.v(v - lo) / t.v(hi - m)) * t.v(t.v(hi - m) / t.v(hi - lo))),
}

type Channel = (t: Trap, c0: number, c1: number, f: number) => number
export const ORIGIN_CHANNELS: Readonly<Record<string, Channel>> = {
  'step-truncate': (t, c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(t.v(f * Math.abs(c1 - c0))),
  'step-truncate-guarded': (t, c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(t.v(f * Math.abs(c1 - c0)) + 1e-11),
  // Written from the maximum (c1) end: the step from c1 rounded up, i.e. toward the midpoint colour c0.
  'max-ceil': (t, c0, c1, f) => c1 - sgn(c1 - c0) * Math.ceil(t.v(t.v(1 - f) * Math.abs(c1 - c0))),
  'max-ceil-guarded': (t, c0, c1, f) => c1 - sgn(c1 - c0) * Math.ceil(t.v(t.v(1 - f) * Math.abs(c1 - c0)) - 1e-11),
  'max-floor': (t, c0, c1, f) => c1 - sgn(c1 - c0) * Math.floor(t.v(t.v(1 - f) * Math.abs(c1 - c0))),
  'max-truncate-signed': (t, c0, c1, f) => c1 + Math.trunc(t.v(t.v(1 - f) * (c0 - c1))),
  'max-value-toward-mid': (t, c0, c1, f) => towardOf(t.v(c1 + t.v(t.v(1 - f) * (c0 - c1))), c0),
  'max-value-toward-max': (t, c0, c1, f) => towardOf(t.v(c1 + t.v(t.v(1 - f) * (c0 - c1))), c1),
  'max-value-floor': (t, c0, c1, f) => Math.floor(t.v(c1 + t.v(t.v(1 - f) * (c0 - c1)))),
  'max-value-round': (t, c0, c1, f) => Math.round(t.v(c1 + t.v(t.v(1 - f) * (c0 - c1)))),
  'value-toward-start': (t, c0, c1, f) => towardOf(t.v(c0 + t.v(f * (c1 - c0))), c0),
  'value-floor': (t, c0, c1, f) => Math.floor(t.v(c0 + t.v(f * (c1 - c0)))),
  'value-truncate': (t, c0, c1, f) => Math.trunc(t.v(c0 + t.v(f * (c1 - c0)))),
  'convex-toward-start': (t, c0, c1, f) => towardOf(t.v(t.v(c0 * t.v(1 - f)) + t.v(c1 * f)), c0),
  'convex-from-max-toward-start': (t, c0, c1, f) => towardOf(t.v(t.v(c1 * f) + t.v(c0 * t.v(1 - f))), c0),
  'convex-floor': (t, c0, c1, f) => Math.floor(t.v(t.v(c0 * t.v(1 - f)) + t.v(c1 * f))),
  // 0..1 channels (c / 255), the lerp in unit space, scaled back by 255.
  'unit-toward-start': (t, c0, c1, f) => { const a = t.v(c0 / 255), b = t.v(c1 / 255); return towardOf(t.v(t.v(a + t.v(f * t.v(b - a))) * 255), c0) },
  'unit-truncate': (t, c0, c1, f) => { const a = t.v(c0 / 255), b = t.v(c1 / 255); return Math.trunc(t.v(t.v(a + t.v(f * t.v(b - a))) * 255)) },
  'unit-convex-truncate': (t, c0, c1, f) => { const a = t.v(c0 / 255), b = t.v(c1 / 255); return Math.trunc(t.v(t.v(t.v(a * t.v(1 - f)) + t.v(b * f)) * 255)) },
  'unit-from-max-truncate': (t, c0, c1, f) => { const a = t.v(c0 / 255), b = t.v(c1 / 255); return Math.trunc(t.v(t.v(b + t.v(t.v(1 - f) * t.v(a - b))) * 255)) },
  'unit-from-max-toward-start': (t, c0, c1, f) => { const a = t.v(c0 / 255), b = t.v(c1 / 255); return towardOf(t.v(t.v(b + t.v(t.v(1 - f) * t.v(a - b))) * 255), c0) },
}

/** Packed-integer channels: the lerp of the whole 0xRRGGBB integer (channels may borrow), truncated toward each end. */
export const ORIGIN_PACKED: Readonly<Record<string, (t: Trap, p0: number, p1: number, f: number) => number>> = {
  'packed-truncate': (t, p0, p1, f) => p0 + Math.trunc(t.v(f * (p1 - p0))),
  'packed-from-max-truncate': (t, p0, p1, f) => p1 + Math.trunc(t.v(t.v(1 - f) * (p0 - p1))),
}

// ---------------------------------------------------------------------------------------------------------------------
// Family `f32` (H3): float32 at every subset of five fraction operations x float32 at the channel's two operations.
// ---------------------------------------------------------------------------------------------------------------------

type MaskedFraction = (t: Trap, r: (on: boolean, x: number) => number, mask: readonly boolean[], v: number, lo: number, hi: number, m: number) => number
/** mask[0] the operands, mask[1..4] the successive operations of each form. */
export const F32_FRACTIONS: Readonly<Record<string, MaskedFraction>> = {
  exact: (t, r, k, v, lo, hi) => { v = r(k[0]!, v); lo = r(k[0]!, lo); hi = r(k[0]!, hi); return r(k[3]!, t.v(r(k[1]!, v - lo) / r(k[2]!, hi - lo))) },
  'from-max': (t, r, k, v, lo, hi) => { v = r(k[0]!, v); lo = r(k[0]!, lo); hi = r(k[0]!, hi); return r(k[4]!, 1 - r(k[3]!, t.v(r(k[1]!, hi - v) / r(k[2]!, hi - lo)))) },
  'g-remap': (t, r, k, v, lo, hi, m) => { v = r(k[0]!, v); lo = r(k[0]!, lo); hi = r(k[0]!, hi); m = r(k[0]!, m); const g = r(k[1]!, t.v((v - m) / (hi - m))), gm = r(k[2]!, t.v((lo - m) / (hi - m))); return r(k[4]!, t.v(r(k[3]!, g - gm) / r(k[3]!, 1 - gm))) },
  'g-reverse': (t, r, k, v, lo, hi, m) => { v = r(k[0]!, v); lo = r(k[0]!, lo); hi = r(k[0]!, hi); m = r(k[0]!, m); const g = r(k[1]!, t.v((hi - v) / (hi - m))), gm = r(k[2]!, t.v((hi - lo) / (hi - m))); return r(k[4]!, 1 - r(k[3]!, t.v(g / gm))) },
}
type MaskedChannel = (t: Trap, r: (on: boolean, x: number) => number, mask: readonly boolean[], c0: number, c1: number, f: number) => number
export const F32_CHANNELS: Readonly<Record<string, MaskedChannel>> = {
  step: (_t, r, k, c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(r(k[1]!, r(k[0]!, f) * Math.abs(c1 - c0))),
  'step-guarded': (_t, r, k, c0, c1, f) => c0 + sgn(c1 - c0) * Math.floor(r(k[1]!, r(k[0]!, f) * Math.abs(c1 - c0)) + 1e-11),
  'max-ceil': (_t, r, k, c0, c1, f) => c1 - sgn(c1 - c0) * Math.ceil(r(k[1]!, r(k[0]!, 1 - f) * Math.abs(c1 - c0))),
  'value-toward-start': (_t, r, k, c0, c1, f) => towardOf(r(k[1]!, c0 + r(k[0]!, f * (c1 - c0))), c0),
  'max-value-toward-mid': (_t, r, k, c0, c1, f) => towardOf(r(k[1]!, c1 + r(k[0]!, (1 - f) * (c0 - c1))), c0),
  'unit-toward-start': (_t, r, k, c0, c1, f) => { const a = r(k[0]!, c0 / 255), b = r(k[0]!, c1 / 255); return towardOf(r(k[1]!, r(k[1]!, a + r(k[1]!, f * r(k[1]!, b - a))) * 255), c0) },
  'unit-truncate': (_t, r, k, c0, c1, f) => { const a = r(k[0]!, c0 / 255), b = r(k[0]!, c1 / 255); return Math.floor(r(k[1]!, r(k[1]!, a + r(k[1]!, f * r(k[1]!, b - a))) * 255)) },
}
const masks = (n: number): boolean[][] => Array.from({ length: 1 << n }, (_, k) => Array.from({ length: n }, (_, i) => ((k >> i) & 1) === 1))
const maskText = (m: readonly boolean[]): string => m.map(b => (b ? '1' : '0')).join('')

// ---------------------------------------------------------------------------------------------------------------------
// Family `precision` (H2): the double representation rounded to a stated precision, at the fraction and at the product.
// ---------------------------------------------------------------------------------------------------------------------

/** Rounds `x` to a `bits`-bit binary significand, to nearest (ties away) or truncated toward zero. */
export function roundBits(x: number, bits: number, mode: 'nearest' | 'truncate'): number {
  if (x === 0 || !Number.isFinite(x)) return x
  const e = Math.floor(Math.log2(Math.abs(x)))
  const scale = 2 ** (bits - 1 - e)
  const y = x * scale
  return (mode === 'nearest' ? Math.round(y) : Math.trunc(y)) / scale
}
/** The precisions: none, float32, p-bit binary (nearest `bin`, truncated `bint`, p = 18..53), n significant decimal digits (6..17). */
export function precisions(): Record<string, (x: number) => number> {
  const out: Record<string, (x: number) => number> = { none: x => x, f32: x => Math.fround(x) }
  for (let p = 18; p <= 53; p++) { out[`bin${p}`] = x => roundBits(x, p, 'nearest'); out[`bint${p}`] = x => roundBits(x, p, 'truncate') }
  for (let n = 6; n <= 17; n++) out[`dec${n}`] = x => Number(x.toPrecision(n))
  return out
}
/** Forms: P rounds the fraction (or the complement r = (hi − v)/(hi − lo)), Q the step product; the step from c0 is returned. */
type PrecisionForm = (v: number, lo: number, hi: number, d: number, P: (x: number) => number, Q: (x: number) => number) => number
export const PRECISION_FORMS: Readonly<Record<string, PrecisionForm>> = {
  'mid-step': (v, lo, hi, d, P, Q) => Math.floor(Q(P((v - lo) / (hi - lo)) * d)),
  'complement-step': (v, lo, hi, d, P, Q) => Math.floor(Q(P(1 - P((hi - v) / (hi - lo))) * d)),
  'complement-raw-step': (v, lo, hi, d, P, Q) => Math.floor(Q((1 - P((hi - v) / (hi - lo))) * d)),
  'max-ceil': (v, lo, hi, d, P, Q) => d - Math.ceil(Q(P((hi - v) / (hi - lo)) * d)),
  'max-ceil-of-complement': (v, lo, hi, d, P, Q) => d - Math.ceil(Q(P(1 - (v - lo) / (hi - lo)) * d)),
  'span-minus-max-step': (v, lo, hi, d, P, Q) => Math.floor(Q(d - P((hi - v) / (hi - lo)) * d)),
}

// ---------------------------------------------------------------------------------------------------------------------
// Family `rational` (H3): exact rationals (BigInt), each operation rounded to a p-bit significand (nearest-even) or exact.
// ---------------------------------------------------------------------------------------------------------------------

interface Q { n: bigint; d: bigint }
/** A finite double as an exact rational (denominator a power of two). */
function qOf(x: number): Q {
  if (!Number.isFinite(x)) throw new Error('qOf: non-finite')
  let d = 1n, y = x
  while (!Number.isInteger(y)) { y *= 2; d *= 2n }
  return { n: BigInt(y), d }
}
const qAdd = (a: Q, b: Q): Q => ({ n: a.n * b.d + b.n * a.d, d: a.d * b.d })
const qSub = (a: Q, b: Q): Q => ({ n: a.n * b.d - b.n * a.d, d: a.d * b.d })
const qMul = (a: Q, b: Q): Q => ({ n: a.n * b.n, d: a.d * b.d })
const qDiv = (a: Q, b: Q): Q => { const n = a.n * b.d, d = a.d * b.n; return d < 0n ? { n: -n, d: -d } : { n, d } }
const qFloor = (a: Q): bigint => { const f = a.n / a.d; return a.n < 0n && f * a.d !== a.n ? f - 1n : f }
const qCeil = (a: Q): bigint => -qFloor({ n: -a.n, d: a.d })
const bitLen = (x: bigint): number => (x === 0n ? 0 : x.toString(2).length)
/** Rounds to a `p`-bit significand, nearest-even (p = Infinity: exact). */
export function qRound(a: Q, p: number): Q {
  if (!Number.isFinite(p) || a.n === 0n) return a
  const neg = a.n < 0n, n = neg ? -a.n : a.n, d = a.d
  let e = BigInt(bitLen(n) - bitLen(d) - (p - 1))
  const scaled = (k: bigint): Q => (k >= 0n ? { n, d: d << k } : { n: n << -k, d })
  let s = scaled(e)
  while (s.n >= s.d << BigInt(p)) { e++; s = scaled(e) }
  while (s.n < s.d << BigInt(p - 1)) { e--; s = scaled(e) }
  let m = s.n / s.d
  const rem = s.n - m * s.d
  if (rem * 2n > s.d || (rem * 2n === s.d && (m & 1n) === 1n)) m++
  const out: Q = e >= 0n ? { n: m << e, d: 1n } : { n: m, d: 1n << -e }
  return neg ? { n: -out.n, d: out.d } : out
}
const ONE: Q = { n: 1n, d: 1n }
type RationalForm = (v: Q, lo: Q, hi: Q, m: Q, c0: number, c1: number, p: number) => bigint
const R = qRound
const fracQ = (v: Q, lo: Q, hi: Q, p: number): Q => R(qDiv(R(qSub(v, lo), p), R(qSub(hi, lo), p)), p)
const compQ = (v: Q, lo: Q, hi: Q, p: number): Q => R(qDiv(R(qSub(hi, v), p), R(qSub(hi, lo), p)), p)
const stepFrom = (c0: number, c1: number, k: bigint): bigint => BigInt(c0) + BigInt(sgn(c1 - c0)) * k
export const RATIONAL_FORMS: Readonly<Record<string, RationalForm>> = {
  step: (v, lo, hi, _m, c0, c1, p) => stepFrom(c0, c1, qFloor(R(qMul(fracQ(v, lo, hi, p), qOf(Math.abs(c1 - c0))), p))),
  'from-max-step': (v, lo, hi, _m, c0, c1, p) => stepFrom(c0, c1, qFloor(R(qMul(R(qSub(ONE, compQ(v, lo, hi, p)), p), qOf(Math.abs(c1 - c0))), p))),
  'from-max-ceil': (v, lo, hi, _m, c0, c1, p) => BigInt(c1) - BigInt(sgn(c1 - c0)) * qCeil(R(qMul(compQ(v, lo, hi, p), qOf(Math.abs(c1 - c0))), p)),
  // FMA: f·(c1 − c0) + c0 rounded once.
  'fma-value-toward-start': (v, lo, hi, _m, c0, c1, p) => { const x = R(qAdd(qMul(fracQ(v, lo, hi, p), qOf(c1 - c0)), qOf(c0)), p); return c1 > c0 ? qFloor(x) : qCeil(x) },
  'fma-from-max-toward-mid': (v, lo, hi, _m, c0, c1, p) => { const x = R(qAdd(qMul(compQ(v, lo, hi, p), qOf(c0 - c1)), qOf(c1)), p); return c1 > c0 ? qFloor(x) : qCeil(x) },
  'fma-from-max-floor': (v, lo, hi, _m, c0, c1, p) => qFloor(R(qAdd(qMul(compQ(v, lo, hi, p), qOf(c0 - c1)), qOf(c1)), p)),
  'fma-convex-toward-start': (v, lo, hi, _m, c0, c1, p) => { const f = fracQ(v, lo, hi, p); const a = R(qMul(R(qSub(ONE, f), p), qOf(c0)), p); const x = R(qAdd(qMul(f, qOf(c1)), a), p); return c1 > c0 ? qFloor(x) : qCeil(x) },
  'fma-convex-from-max-toward-start': (v, lo, hi, _m, c0, c1, p) => { const r = compQ(v, lo, hi, p); const a = R(qMul(R(qSub(ONE, r), p), qOf(c1)), p); const x = R(qAdd(qMul(r, qOf(c0)), a), p); return c1 > c0 ? qFloor(x) : qCeil(x) },
  'g-remap-step': (v, lo, hi, m, c0, c1, p) => { const g = R(qDiv(R(qSub(v, m), p), R(qSub(hi, m), p)), p), gm = R(qDiv(R(qSub(lo, m), p), R(qSub(hi, m), p)), p); const f = R(qDiv(R(qSub(g, gm), p), R(qSub(ONE, gm), p)), p); return stepFrom(c0, c1, qFloor(R(qMul(f, qOf(Math.abs(c1 - c0))), p))) },
  'g-remap-fma-value-toward-start': (v, lo, hi, m, c0, c1, p) => { const g = R(qDiv(R(qSub(v, m), p), R(qSub(hi, m), p)), p), gm = R(qDiv(R(qSub(lo, m), p), R(qSub(hi, m), p)), p); const f = R(qDiv(R(qSub(g, gm), p), R(qSub(ONE, gm), p)), p); const x = R(qAdd(qMul(f, qOf(c1 - c0)), qOf(c0)), p); return c1 > c0 ? qFloor(x) : qCeil(x) },
}
export const RATIONAL_PRECISIONS: readonly number[] = [24, 53, 64, Number.POSITIVE_INFINITY]

// ---------------------------------------------------------------------------------------------------------------------
// The candidates and their evaluation.
// ---------------------------------------------------------------------------------------------------------------------

/** Per-channel evaluation inside the segment: endpoints and equal channels keep production's colour. */
function perChannel(s: CfScaleSegment, ch: (c0: number, c1: number) => number | bigint | undefined): number | undefined {
  const a = channelsOf(s.c0), b = channelsOf(s.c1)
  const out: number[] = []
  for (let i = 0; i < 3; i++) {
    if (a[i] === b[i]) { out.push(a[i]!); continue }
    const c = ch(a[i]!, b[i]!)
    if (c === undefined) return undefined
    out.push(Number(c))
  }
  return pack(out)
}
/** Strictly inside the segment (endpoints draw their colours under every formulation). */
const interior = (s: CfScaleSegment): boolean => s.value > s.lo && s.value < s.hi

function originCandidates(): UpperWholeCandidate[] {
  const out: UpperWholeCandidate[] = []
  for (const [fid, frac] of Object.entries(ORIGIN_FRACTIONS)) {
    for (const [cid, ch] of Object.entries(ORIGIN_CHANNELS)) out.push({
      id: `origin/${fid}/${cid}`, family: 'origin',
      evaluate: s => {
        const x = inputsOf(s); if (!x) return undefined
        const t = new Trap(), f = clamp01(frac(t, x.v, x.lo, x.hi, x.m))
        const argb = perChannel(s, (c0, c1) => ch(t, c0, c1, f))
        return t.bad ? undefined : argb
      },
    })
    for (const [cid, ch] of Object.entries(ORIGIN_PACKED)) out.push({
      id: `origin/${fid}/${cid}`, family: 'origin',
      evaluate: s => {
        const x = inputsOf(s); if (!x) return undefined
        const t = new Trap(), f = clamp01(frac(t, x.v, x.lo, x.hi, x.m))
        const p = ch(t, s.c0 & 0xffffff, s.c1 & 0xffffff, f)
        if (t.bad || !Number.isInteger(p) || p < 0 || p > 0xffffff) return undefined
        return (0xff000000 | p) >>> 0
      },
    })
  }
  return out
}

function f32Candidates(): UpperWholeCandidate[] {
  const out: UpperWholeCandidate[] = []
  for (const [fid, frac] of Object.entries(F32_FRACTIONS)) for (const fm of masks(5)) for (const [cid, ch] of Object.entries(F32_CHANNELS)) for (const cm of masks(2)) out.push({
    id: `f32/${fid}[${maskText(fm)}]/${cid}[${maskText(cm)}]`, family: 'f32',
    evaluate: s => {
      const x = inputsOf(s); if (!x) return undefined
      const t = new Trap()
      const r = (on: boolean, y: number): number => (on ? t.f(y) : t.v(y))
      const f = clamp01(frac(t, r, fm, x.v, x.lo, x.hi, x.m))
      const argb = perChannel(s, (c0, c1) => ch(t, r, cm, c0, c1, f))
      return t.bad ? undefined : argb
    },
  })
  return out
}

function precisionCandidates(): UpperWholeCandidate[] {
  const P = precisions()
  const out: UpperWholeCandidate[] = []
  for (const [form, fn] of Object.entries(PRECISION_FORMS)) for (const [pk, p] of Object.entries(P)) for (const [qk, q] of Object.entries(P)) out.push({
    id: `precision/${form}/${pk}/${qk}`, family: 'precision',
    evaluate: s => {
      const x = inputsOf(s); if (!x) return undefined
      const t = new Trap()
      const pp = (y: number): number => t.v(p(t.v(y))), qq = (y: number): number => t.v(q(t.v(y)))
      return perChannel(s, (c0, c1) => {
        const d = Math.abs(c1 - c0)
        const k = Math.max(0, Math.min(d, fn(x.v, x.lo, x.hi, d, pp, qq)))
        return t.bad ? undefined : c0 + sgn(c1 - c0) * k
      })
    },
  })
  return out
}

function rationalCandidates(): UpperWholeCandidate[] {
  const out: UpperWholeCandidate[] = []
  for (const [form, fn] of Object.entries(RATIONAL_FORMS)) for (const p of RATIONAL_PRECISIONS) out.push({
    id: `rational/${form}/${Number.isFinite(p) ? `p${p}` : 'exact'}`, family: 'rational',
    evaluate: s => {
      const x = inputsOf(s); if (!x) return undefined
      const [v, lo, hi, m] = [qOf(x.v), qOf(x.lo), qOf(x.hi), qOf(x.m)]
      return perChannel(s, (c0, c1) => {
        const lohi = c0 < c1 ? [BigInt(c0), BigInt(c1)] : [BigInt(c1), BigInt(c0)]
        const c = fn(v, lo, hi, m, c0, c1, p)
        return c < lohi[0]! ? lohi[0]! : c > lohi[1]! ? lohi[1]! : c
      })
    },
  })
  return out
}

let built: UpperWholeCandidate[] | undefined
/** Every candidate of the extended search, in enumeration order (built on first use). */
export function upperWholeCandidates(): readonly UpperWholeCandidate[] {
  return (built ??= [...originCandidates(), ...f32Candidates(), ...precisionCandidates(), ...rationalCandidates()])
}

/** Where a candidate replaces production's arithmetic (inside the segment); elsewhere production's truncation stands. */
export const UPPER_WHOLE_SCOPES: Readonly<Record<string, (s: CfScaleSegment) => boolean>> = {
  'upper-3-colour': s => s.colours === 3 && s.segment === 1,
  '3-colour': s => s.colours === 3,
  all: () => true,
}
/** Production's truncation (the baseline, and every out-of-scope cell). */
export const truncationOf = (s: CfScaleSegment): number => interpolateScaleColor(s.c0, s.c1, s.fraction)

export interface UpperWholeResult { id: string; family: UpperWholeCandidate['family']; scope: string; inScope: number; mismatches: string[] }

/** Evaluates candidates at every scope; a cell's key is `caseId:ref`. Out-of-scope cells and endpoints keep truncation. */
export function evaluateUpperWhole(cells: readonly SegmentCell[], candidates: readonly UpperWholeCandidate[] = upperWholeCandidates()): UpperWholeResult[] {
  const key = (c: SegmentCell): string => `${c.caseId}:${c.ref}`
  const out: UpperWholeResult[] = []
  for (const [scope, inScope] of Object.entries(UPPER_WHOLE_SCOPES)) {
    const scoped = cells.filter(c => inScope(c.segment) && interior(c.segment))
    // Cells outside the scope are drawn by truncation under every candidate: their mismatches are shared.
    const outside = cells.filter(c => !(inScope(c.segment) && interior(c.segment))).filter(c => hex(truncationOf(c.segment)) !== c.excel).map(key)
    for (const cand of candidates) {
      const mismatches = [...outside]
      for (const c of scoped) {
        const got = cand.evaluate(c.segment)
        if (got === undefined || hex(got) !== c.excel) mismatches.push(key(c))
      }
      out.push({ id: cand.id, family: cand.family, scope, inScope: scoped.length, mismatches })
    }
  }
  return out
}

/** The baseline: cells where production's truncation is not Excel's fill. */
export function truncationMisses(cells: readonly SegmentCell[]): string[] {
  return cells.filter(c => hex(truncationOf(c.segment)) !== c.excel).map(c => `${c.caseId}:${c.ref}`)
}

/**
 * The capture that followed this search (stage2-scale-upper-whole, 54fc3388, exploratory): `{ include: 'pre-capture' }`
 * leaves it out, reproducing the numbers of recipe §9.1 as they were when the case was designed; the default includes it.
 */
export const UPPER_WHOLE_POST_CAPTURE_CASES: readonly string[] = [S2_UPPER_WHOLE_ID]
/**
 * Captures after this search's own follow-up, left out in every mode: the search is recorded as of 54fc3388 (recipe §9.5).
 * stage2-scale-upper-window (ed0fccb6) is searched by scripts/cf-refs/upper-window-candidates.ts, which includes it;
 * stage2-scale-raw-floor (ce3f701d) by scripts/cf-refs/raw-floor-census.ts.
 */
export const UPPER_WHOLE_LATER_CASES: readonly string[] = ['stage2-scale-upper-window', 'stage2-scale-raw-floor']

export async function upperWholeCells(rootDir: string = REPO_ROOT, opts: { include?: 'pre-capture' | 'all' } = {}): Promise<SegmentCell[]> {
  const manifest = JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const { cells, faults } = await collectSegmentCells(manifest, rootDir, { include: 'all' })
  if (faults.length > 0) throw new Error(`upper-whole: harness faults ${faults.join('; ')}`)
  const recorded = cells.filter(c => !UPPER_WHOLE_LATER_CASES.includes(c.caseId))
  return (opts.include ?? 'all') === 'all' ? recorded : recorded.filter(c => !UPPER_WHOLE_POST_CAPTURE_CASES.includes(c.caseId))
}

// ---------------------------------------------------------------------------------------------------------------------
// The exploratory case's predictions (docs/validation/xlsx-cf-upper-whole-predictions.md, `--write-predictions`).
// ---------------------------------------------------------------------------------------------------------------------

/** The float32 near miss of §9.1 (never a reading): float32 from the maximum, float32 product, upper 3-colour scope only. */
export const NEAR_MISS_ID = 'f32/from-max[01111]/step[11]'

/** An exact decimal as a rational (BigInt numerator, power-of-ten denominator). */
function decimalQ(text: string): Q {
  const m = /^(-?)(\d+)(?:\.(\d+))?$/.exec(text)
  if (!m) throw new Error(`upper-whole: ${text} is not a plain decimal`)
  const frac = m[3] ?? ''
  const n = BigInt(`${m[2]}${frac}`) * (m[1] ? -1n : 1n)
  return { n, d: 10n ** BigInt(frac.length) }
}
const gcd = (a: bigint, b: bigint): bigint => { a = a < 0n ? -a : a; while (b) [a, b] = [b, a % b]; return a }

export interface UpperWholePrediction {
  ref: string
  configuration: string
  block: string
  value: string
  /** The exact fraction p/q of the segment (from the decimals). */
  fraction: string
  /** The segment: L (lower or only) or U (upper). */
  segment: 'L' | 'U'
  /** Per RGB channel: d = c1 − c0, and whether f·|d| is whole in exact arithmetic. */
  channels: Array<{ d: number; whole: boolean }>
  truncate: string
  /** Truncation with every exactly-whole channel's step one unit smaller (the 377564e0 departure, per channel). */
  minusOne: string
  nearMiss: string
  /** Production today: `painted` or the withhold reason. */
  production: string
  /** The truncate token (twin or literal). */
  token: string
}

/** Every designated cell's predictions: truncate (the reading), truncation minus one step and the near miss (neither a reading). */
export function upperWholePredictions(): UpperWholePrediction[] {
  const near = upperWholeCandidates().find(c => c.id === NEAR_MISS_ID)!
  const { configurations } = s2UpperWholeLayout()
  const tokenOf = new Map(configurations.flatMap(cfg => Object.entries(cfg.hypotheses.truncate!)))
  return s2UpperWholeCells().filter(c => c.designated).map((c: S2UpperWholeCell): UpperWholePrediction => {
    const tIdx = c.segment
    const [v, lo, hi] = [decimalQ(c.text), decimalQ(c.block.thresholds[tIdx]!), decimalQ(c.block.thresholds[tIdx + 1]!)]
    const f = qDiv(qSub(v, lo), qSub(hi, lo))
    const g = gcd(f.n, f.d)
    const a = channelsOf(Number.parseInt(c.from, 16) >>> 0), b = channelsOf(Number.parseInt(c.to, 16) >>> 0)
    const interiorCell = f.n > 0n && f.n < f.d
    const channels = [0, 1, 2].map(i => { const d = b[i]! - a[i]!; return { d, whole: d !== 0 && interiorCell && (f.n * BigInt(Math.abs(d))) % f.d === 0n } })
    const truncate = s2UpperWholeTruncation(c)
    const tr = channelsOf(Number.parseInt(truncate, 16) >>> 0)
    const minusOne = hex(pack(tr.map((x, i) => (channels[i]!.whole ? x - sgn(channels[i]!.d) : x)))!)
    const t = c.block.thresholds.map(Number)
    const seg: CfScaleSegment = { value: c.value, lo: c.lo, hi: c.hi, fraction: clamp01((c.value - c.lo) / (c.hi - c.lo)), c0: Number.parseInt(c.from, 16) >>> 0, c1: Number.parseInt(c.to, 16) >>> 0, colours: c.block.colors.length as 2 | 3, segment: c.segment, loKind: 'num', hiKind: 'num', thresholds: t }
    const nearArgb = seg.colours === 3 && seg.segment === 1 && interior(seg) ? near.evaluate(seg) : truncationOf(seg)
    const kinds = c.block.cfvo.map(x => x.type)
    const point = (k: number): CfScalePoint => (kinds[k] === 'num' ? { bound: { type: 'num', value: t[k]! }, argb: 0 } : { bound: { type: kinds[k] as 'min' }, argb: 0 })
    const production = seg.colours === 3 && numSegmentUnmeasured(point(tIdx), point(tIdx + 1), seg.value, seg.lo, seg.hi, seg.fraction) ? 'withheld (num-segment)'
      : seg.colours === 3 && seg.segment === 1 && upperSegmentWholeProduct(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction) ? 'withheld (upper whole product)' : 'painted'
    return { ref: c.ref, configuration: c.configuration!, block: c.block.id, value: c.text, fraction: `${f.n / g}/${f.d / g}`, segment: c.segment === 1 ? 'U' : 'L', channels, truncate, minusOne, nearMiss: nearArgb === undefined ? 'undefined' : hex(nearArgb), production, token: tokenOf.get(c.ref)! }
  })
}

function writePredictions(): void {
  const rows = upperWholePredictions()
  const ch = (r: UpperWholePrediction): string => r.channels.map(x => `${x.d >= 0 ? '+' : ''}${x.d}${x.whole ? '*' : ''}`).join(' ')
  const lines = [
    `# ${S2_UPPER_WHOLE_ID}: predictions at every designated cell (exploratory)`,
    '',
    'Generated by `npx tsx scripts/cf-refs/upper-whole-candidates.ts --write-predictions` from the catalogue. **Only `truncate` is a reading** (the control); the other two columns are for interpreting the capture and can never be registered from it:',
    '- **minus one**: truncation with the step one unit smaller on every channel whose exact product f·|d| is whole (the 377564e0 departure, applied per channel);',
    `- **near miss**: \`${NEAR_MISS_ID}\` (float32 from the maximum), contradicted by E4 of stage2-scale-thresholds (§9.1), upper 3-colour segments only.`,
    '',
    'Channels list d = c1 − c0 per R, G, B; `*` marks an exactly-whole product. Production is today\'s default (`painted` at truncation, or the withhold).',
    '',
    '| Cell | Config | Block | Value | Seg | f | d (R G B) | truncate (token) | minus one | near miss | Production |',
    '|---|---|---|---|---|---|---|---|---|---|---|',
    ...rows.map(r => `| ${r.ref} | ${r.configuration} | ${r.block} | ${r.value} | ${r.segment} | ${r.fraction} | ${ch(r)} | ${r.truncate}${r.token.startsWith('twin:') ? ` (${r.token})` : ''} | ${r.minusOne === r.truncate ? '=' : r.minusOne} | ${r.nearMiss === r.truncate ? '=' : r.nearMiss} | ${r.production} |`),
    '',
  ]
  writeFileSync(resolve(REPO_ROOT, 'docs/validation/xlsx-cf-upper-whole-predictions.md'), lines.join('\n'))
}

async function main(): Promise<void> {
  if (process.argv.includes('--write-predictions')) { writePredictions(); return }
  const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const collected = await collectSegmentCells(manifest, REPO_ROOT, { include: 'all' })
  // `--pre-capture`: the cells as they were before stage2-scale-upper-whole was captured (recipe §9.1).
  const keep = (k: { caseId: string }): boolean => !UPPER_WHOLE_LATER_CASES.includes(k.caseId) && (!process.argv.includes('--pre-capture') || !UPPER_WHOLE_POST_CAPTURE_CASES.includes(k.caseId))
  const [cells, excluded, faults] = [collected.cells.filter(keep), collected.excluded.filter(keep), collected.faults]
  const short = (k: string): string => k.replace(/^stage\d-(scale-)?/, '')
  console.log(`## Cells\n\n${cells.length} captured colour-scale cells; ${excluded.filter(e => e.group === 'a').length} excluded as inapplicable (group a), ${excluded.filter(e => e.group === 'b').length} dropped by a production gate (group b, ${excluded.filter(e => e.group === 'b' && e.threeColour).length} under a 3-colour scale); ${faults.length} harness faults.`)
  const base = truncationMisses(cells)
  console.log(`\nBaseline (production's truncation): ${base.length} mismatches: ${base.map(short).join(', ')}.`)
  const results = evaluateUpperWhole(cells)
  const families = ['origin', 'f32', 'precision', 'rational'] as const
  console.log('\n## Survivors\n\n| Family | Scope | Candidates | In-scope cells | Survivors | Fewest mismatches | Best candidates (first 3) and their mismatches |\n|---|---|---|---|---|---|---|')
  for (const fam of families) for (const scope of Object.keys(UPPER_WHOLE_SCOPES)) {
    const rs = results.filter(r => r.family === fam && r.scope === scope)
    const best = Math.min(...rs.map(r => r.mismatches.length))
    const top = rs.filter(r => r.mismatches.length === best)
    console.log(`| ${fam} | ${scope} | ${rs.length} | ${rs[0]!.inScope} | ${rs.filter(r => r.mismatches.length === 0).length} | ${best} (${top.length} candidates) | ${top.slice(0, 3).map(r => `\`${r.id}\``).join(', ')}: ${top[0]!.mismatches.slice(0, 6).map(short).join(', ')}${top[0]!.mismatches.length > 6 ? ', …' : ''} |`)
  }
  const near = results.filter(r => r.mismatches.length > 0 && r.mismatches.length <= 2)
  const bySet = new Map<string, UpperWholeResult[]>()
  for (const r of near) { const k = `${r.scope}|${r.mismatches.join(',')}`; bySet.set(k, [...(bySet.get(k) ?? []), r]) }
  console.log(`\n## Near misses (1 or 2 mismatches): ${near.length} candidate x scope pairs\n\n| Scope | Mismatches | Candidates | Examples |\n|---|---|---|---|`)
  for (const [k, rs] of [...bySet].sort((a, b) => a[1][0]!.mismatches.length - b[1][0]!.mismatches.length)) console.log(`| ${k.split('|')[0]} | ${rs[0]!.mismatches.map(short).join(', ')} | ${rs.length} | ${rs.slice(0, 3).map(r => `\`${r.id}\``).join(', ')} |`)
  console.log(`\nTotal: ${results.length / Object.keys(UPPER_WHOLE_SCOPES).length} candidates x ${Object.keys(UPPER_WHOLE_SCOPES).length} scopes; survivors at any scope: ${results.filter(r => r.mismatches.length === 0).length}.`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
