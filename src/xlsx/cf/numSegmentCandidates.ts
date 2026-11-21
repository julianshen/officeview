/**
 * The num-segment follow-up: every closed candidate formulation of the colour inside a colour-scale segment, built lazily by
 * id (nothing is constructed until asked). Shared by the candidate search (scripts/cf-refs/num-segment-candidates.ts) and
 * production's class-consensus evaluation (`colorScale.ts`, `numSegmentConsensusColor`), so a class member's colour is
 * one implementation everywhere. Ids: `<fraction>/<channel>`, `ratio-step-truncate[-unguarded]`, `int<N>-<form>`,
 * `sweep-fixed-<N>`, `sweep-roundtrip[-reciprocal|-after]-<K>`.
 */
import type { CfScaleSegment } from './colorScale'

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
const SWEEPS = (): Readonly<Record<string, (f: number, d: number) => number>> => ({
  ...Object.fromEntries(SWEEP_N.map(N => [`sweep-fixed-${N}`, (f: number, d: number) => Math.floor((Math.floor(f * N) / N) * d)])),
  ...Object.fromEntries(SWEEP_K.flatMap(K => [
    [`sweep-roundtrip-${K}`, (f: number, d: number) => Math.floor(((f * K) / K) * d)],
    [`sweep-roundtrip-reciprocal-${K}`, (f: number, d: number) => Math.floor(f * K * (1 / K) * d)],
    [`sweep-roundtrip-after-${K}`, (f: number, d: number) => Math.floor((f * K * d) / K)],
  ])),
})

export interface Candidate { id: string; colour: (s: CfScaleSegment) => number }

const channelsOf = (argb: number): number[] => [(argb >>> 24) & 0xff, (argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]
function composeArgb(s: CfScaleSegment, channel: (c0: number, c1: number) => number): number {
  const a = channelsOf(s.c0), b = channelsOf(s.c1)
  // Alpha is never observed (DisplayFormat fills are opaque); every captured endpoint is FF.
  const out = [a[0]!, ...[1, 2, 3].map(i => channel(a[i]!, b[i]!))]
  return ((out[0]! << 24) | (out[1]! << 16) | (out[2]! << 8) | out[3]!) >>> 0
}

/** Every candidate: fraction × channel, plus the fixed-point family over the exact fraction. */
function buildCandidates(): Candidate[] {
  return [
  ...Object.entries(FRACTIONS).flatMap(([fid, frac]) => Object.entries(CHANNELS)
    // The ratio channels read the segment directly (no fraction): only with the exact fraction, once.
    .filter(([cid]) => !cid.startsWith('ratio-') || fid === 'exact')
    .map(([cid, ch]) => ({ id: cid.startsWith('ratio-') ? cid : `${fid}/${cid}`, colour: (s: CfScaleSegment) => { const f = clamp01(frac(s)); return composeArgb(s, (c0, c1) => ch(c0, c1, f, s)) } }))),
  ...Object.entries(FIXED).map(([id, ch]) => ({ id, colour: (s: CfScaleSegment) => { const f = clamp01(FRACTIONS.exact!(s)); return composeArgb(s, (c0, c1) => ch(c0, c1, f)) } })),
  ...Object.entries(SWEEPS()).map(([id, step]) => ({ id, colour: (s: CfScaleSegment) => { const f = clamp01(FRACTIONS.exact!(s)); return composeArgb(s, (c0, c1) => c0 + sgn(c1 - c0) * step(f, Math.abs(c1 - c0))) } })),
  ]
}
let built: Candidate[] | undefined
let byId: Map<string, Candidate> | undefined
/** Every candidate, in the search's enumeration order (built on first use). */
export function numSegmentCandidates(): readonly Candidate[] {
  return (built ??= buildCandidates())
}
/** One candidate by id (undefined: no such candidate). */
export function numSegmentCandidate(id: string): Candidate | undefined {
  return (byId ??= new Map(numSegmentCandidates().map(c => [c.id, c]))).get(id)
}
