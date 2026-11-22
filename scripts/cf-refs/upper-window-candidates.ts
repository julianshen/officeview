/**
 * The upper-segment window follow-up (brief cf-upperwindow, Phase 1 task 1): the census of every captured 3-colour
 * upper-segment interior cell whose step product lies inside production's 1e-11 truncation-guard window without being an
 * IEEE integer, and a candidate search for a closed formula that separates the cells Excel moved there from those it did not.
 *
 * PR #99 widened the upper-segment whole-product withhold (`upperSegmentWholeProduct`, `cf-unmeasured-upper-segment-whole-
 * product`) from `Number.isInteger(f·d)` to the window |f·d − k| <= 1e-11 (0 < k < d), a fail-closed ruling from literal-num
 * evidence (stage2-scale-upper-whole F12, F15, F16, F18, F19 moved; F8, F11, F13 did not). The only statistical-bound cells
 * inside the window are stage2-scale-thresholds E4 and its replica D19 (both drawn at guarded truncation), so the window's
 * withhold of statistical shapes is unmeasured. This search asks whether any closed deterministic formula explains the
 * window cells while staying exact on every other captured cell that a reading could paint.
 *
 * Scoring: the cells of `collectSegmentCells` (every measured case, both follow-up captures included), minus the cells that
 * stay withheld whatever a window reading says: upper-segment interior cells of a 3-colour scale with an IEEE-integer whole
 * product on some channel (the integer shape the window reading never narrows). A candidate replaces production's
 * arithmetic strictly inside one of four scopes; elsewhere production's guarded truncation stands:
 * - `window`: the upper-segment window cells only (the cells a window reading paints);
 * - `upper`: every upper-segment interior cell of a 3-colour scale;
 * - `3-colour`: every 3-colour segment;
 * - `all`: every segment.
 * A candidate SURVIVES a scope when it reproduces Excel's DisplayFormat fill at every scored cell. Evaluation fails closed:
 * any non-finite intermediate, or a channel that is not a whole number in 0..255, is a mismatch.
 *
 * Families: the 48,345 candidates of the extended search (`upper-whole-candidates.ts`), plus
 * - `form/…`: the step product computed in another operation order (13 forms: product before division, reciprocal
 *   multiplies, from the maximum, split quotients, percent round trips), with the differences v − lo, hi − lo and hi − v
 *   optionally snapped to 15, 16 or 17 significant decimals first, and the step rounded by floor, guarded floor, floor of the
 *   15-digit product, or strictly below the product (ceil − 1);
 * - `threshold/…`: Excel's statistical threshold modelled as another double: each statistical bound (min, max, percent,
 *   percentile or a formula-resolved bound; never a literal num) moved by up to two ulps either way, with guarded or raw
 *   floor steps. This covers every PERCENTILE or percent arithmetic whose result lands within two ulps of production's.
 * Nothing here is registered or reachable from production (brief hard rule).
 *
 * Run: `npx tsx scripts/cf-refs/upper-window-candidates.ts` (docs/validation/xlsx-cf-recipes-stage2.md §10.1);
 * `--write-predictions` writes docs/validation/xlsx-cf-upper-window-predictions.md from the catalogue (§10.3).
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { CF_TRUNCATION_GUARD, type CfScaleSegment } from '../../src/xlsx/cf/colorScale'
import { S2_UPPER_WINDOW_ID, S2_UPPER_WINDOW_READINGS, s2UpperWindowCells, s2UpperWindowClass, s2UpperWindowLayout, s2UpperWindowPrediction, type CfManifest, type S2UpperWindowCell } from './cases'
import { collectSegmentCells, type SegmentCell } from './num-segment-candidates'
import { truncationOf, upperWholeCandidates } from './upper-whole-candidates'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

const hex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0')
const channelsOf = (argb: number): [number, number, number] => [(argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]
const sgn = (d: number): number => (d >= 0 ? 1 : -1)
const interior = (s: CfScaleSegment): boolean => s.value > s.lo && s.value < s.hi
const upper3 = (s: CfScaleSegment): boolean => s.colours === 3 && s.segment === 1

/** One channel's product class at production's fraction: `integer` (an IEEE integer k, 0 < k < d), `below` / `above` (inside the window, not an integer) or `none`. */
export type ProductClass = 'integer' | 'below' | 'above' | 'none'
export function productClass(f: number, d: number): ProductClass {
  if (d === 0) return 'none'
  const p = f * d, k = Math.round(p)
  if (!(k > 0 && k < d)) return 'none'
  if (Number.isInteger(p)) return 'integer'
  if (Math.abs(p - k) <= CF_TRUNCATION_GUARD) return p < k ? 'below' : 'above'
  return 'none'
}
export const channelClasses = (s: CfScaleSegment): ProductClass[] => { const a = channelsOf(s.c0), b = channelsOf(s.c1); return [0, 1, 2].map(i => productClass(s.fraction, Math.abs(b[i]! - a[i]!))) }
/** Upper-segment interior of a 3-colour scale with an IEEE-integer whole product on some channel: withheld whatever a window reading says. */
export const integerShape = (s: CfScaleSegment): boolean => upper3(s) && interior(s) && channelClasses(s).includes('integer')
/** The window shape: upper-segment interior of a 3-colour scale, no IEEE-integer channel, some channel inside the window. */
export const windowShape = (s: CfScaleSegment): boolean => upper3(s) && interior(s) && !channelClasses(s).includes('integer') && channelClasses(s).some(c => c === 'below' || c === 'above')
/** A segment bound by two literal num cfvo (the num-segment shape's bounds). */
export const literalBounds = (s: CfScaleSegment): boolean => s.loKind === 'num' && s.hiKind === 'num'

/** The step a channel takes from c0 under a rule, per channel, endpoints and equal channels at production's colour; undefined fails closed. */
function perChannel(s: CfScaleSegment, step: (d: number) => number | undefined): number | undefined {
  const a = channelsOf(s.c0), b = channelsOf(s.c1)
  const out: number[] = []
  for (let i = 0; i < 3; i++) {
    const d = Math.abs(b[i]! - a[i]!)
    if (d === 0) { out.push(a[i]!); continue }
    const k = step(d)
    if (k === undefined || !Number.isFinite(k) || !Number.isInteger(k)) return undefined
    out.push(a[i]! + sgn(b[i]! - a[i]!) * Math.max(0, Math.min(d, k)))
  }
  if (!out.every(c => Number.isInteger(c) && c >= 0 && c <= 255)) return undefined
  return ((0xff << 24) | (out[0]! << 16) | (out[1]! << 8) | out[2]!) >>> 0
}

export interface UpperWindowCandidate {
  id: string
  family: 'origin' | 'f32' | 'precision' | 'rational' | 'form' | 'threshold' | 'reading'
  evaluate: (s: CfScaleSegment) => number | undefined
}

// ---------------------------------------------------------------------------------------------------------------------
// Family `form`: the step product in another operation order, optional decimal snapping of the differences.
// ---------------------------------------------------------------------------------------------------------------------

const decimal = (n: number) => (x: number): number => Number(x.toPrecision(n))
export const FORM_SNAPS: Readonly<Record<string, (x: number) => number>> = { none: x => x, dec15: decimal(15), dec16: decimal(16), dec17: decimal(17) }
type Form = (v: number, lo: number, hi: number, d: number, snap: (x: number) => number) => number
export const FORMS: Readonly<Record<string, Form>> = {
  'quotient-times-d': (v, lo, hi, d, s) => (s(v - lo) / s(hi - lo)) * d,
  'product-over-span': (v, lo, hi, d, s) => (s(v - lo) * d) / s(hi - lo),
  'over-span-per-unit': (v, lo, hi, d, s) => s(v - lo) / (s(hi - lo) / d),
  'times-d-over-span': (v, lo, hi, d, s) => s(v - lo) * (d / s(hi - lo)),
  'reciprocal-span': (v, lo, hi, d, s) => s(v - lo) * (1 / s(hi - lo)) * d,
  'from-max': (v, lo, hi, d, s) => (1 - s(hi - v) / s(hi - lo)) * d,
  'd-minus-max-quotient': (v, lo, hi, d, s) => d - (s(hi - v) / s(hi - lo)) * d,
  'd-minus-max-product': (v, lo, hi, d, s) => d - (s(hi - v) * d) / s(hi - lo),
  'split-quotients': (v, lo, hi, d, s) => (v / s(hi - lo) - lo / s(hi - lo)) * d,
  'split-products': (v, lo, hi, d, s) => (v * d - lo * d) / s(hi - lo),
  'percent-roundtrip': (v, lo, hi, d, s) => (((s(v - lo) / s(hi - lo)) * 100) / 100) * d,
  'percent-product': (v, lo, hi, d, s) => ((s(v - lo) * 100) / s(hi - lo)) * d / 100,
  'dec15-fraction': (v, lo, hi, d, s) => decimal(15)(s(v - lo) / s(hi - lo)) * d,
}
export const FORM_ROUNDINGS: Readonly<Record<string, (p: number) => number>> = {
  floor: p => Math.floor(p),
  guarded: p => Math.floor(p + CF_TRUNCATION_GUARD),
  'dec15-floor': p => Math.floor(decimal(15)(p)),
  'strictly-below': p => Math.ceil(p) - 1,
}

function formCandidates(): UpperWindowCandidate[] {
  const out: UpperWindowCandidate[] = []
  for (const [sn, snap] of Object.entries(FORM_SNAPS)) for (const [fn, form] of Object.entries(FORMS)) for (const [rn, round] of Object.entries(FORM_ROUNDINGS)) out.push({
    id: `form/${sn}/${fn}/${rn}`, family: 'form',
    evaluate: s => {
      if (![s.value, s.lo, s.hi].every(Number.isFinite) || s.hi - s.lo === 0) return undefined
      return perChannel(s, d => { const p = form(s.value, s.lo, s.hi, d, snap); return Number.isFinite(p) ? round(p) : undefined })
    },
  })
  return out
}

// ---------------------------------------------------------------------------------------------------------------------
// Family `threshold`: a statistical bound's double moved by up to two ulps (Excel's own PERCENTILE / percent arithmetic).
// ---------------------------------------------------------------------------------------------------------------------

const f64 = new Float64Array(1), i64 = new BigInt64Array(f64.buffer)
/** The double `n` ulps from `x` (n may be negative); finite non-zero x only. */
export function ulpStep(x: number, n: number): number {
  if (n === 0 || !Number.isFinite(x) || x === 0) return x
  f64[0] = x
  i64[0] = i64[0]! + BigInt(x > 0 ? n : -n)
  return f64[0]!
}
const statistical = (kind: CfScaleSegment['loKind']): boolean => kind !== 'num'

function thresholdCandidates(): UpperWindowCandidate[] {
  const out: UpperWindowCandidate[] = []
  for (let a = -2; a <= 2; a++) for (let b = -2; b <= 2; b++) for (const [rn, round] of [['guarded', FORM_ROUNDINGS.guarded!], ['floor', FORM_ROUNDINGS.floor!]] as const) {
    if (a === 0 && b === 0) continue
    out.push({
      id: `threshold/lo${a >= 0 ? '+' : ''}${a}/hi${b >= 0 ? '+' : ''}${b}/${rn}`, family: 'threshold',
      evaluate: s => {
        const lo = statistical(s.loKind) ? ulpStep(s.lo, a) : s.lo, hi = statistical(s.hiKind) ? ulpStep(s.hi, b) : s.hi
        if (![s.value, lo, hi].every(Number.isFinite) || !(hi > lo)) return undefined
        const f = Math.max(0, Math.min(1, (s.value - lo) / (hi - lo)))
        if (!Number.isFinite(f)) return undefined
        return perChannel(s, d => round(f * d))
      },
    })
  }
  return out
}

/** The two readings the case pre-registers whatever the search finds (brief decisions): guarded truncation and raw floor. */
export const READING_CANDIDATES: readonly UpperWindowCandidate[] = [
  { id: 'reading/guarded', family: 'reading', evaluate: s => truncationOf(s) },
  { id: 'reading/raw-floor', family: 'reading', evaluate: s => perChannel(s, d => { const p = s.fraction * d; return Number.isFinite(p) ? Math.floor(p) : undefined }) },
]

let built: UpperWindowCandidate[] | undefined
/** Every candidate: the extended search's, then `form`, `threshold` and the two readings. */
export function upperWindowCandidates(): readonly UpperWindowCandidate[] {
  return (built ??= [...upperWholeCandidates(), ...formCandidates(), ...thresholdCandidates(), ...READING_CANDIDATES])
}

export const UPPER_WINDOW_SCOPES: Readonly<Record<string, (s: CfScaleSegment) => boolean>> = {
  window: windowShape,
  upper: s => upper3(s) && interior(s),
  '3-colour': s => s.colours === 3 && interior(s),
  all: s => interior(s),
}

export interface UpperWindowResult { id: string; family: UpperWindowCandidate['family']; scope: string; inScope: number; mismatches: string[] }
const keyOf = (c: SegmentCell): string => `${c.caseId}:${c.ref}`

/** The scored cells: every collected cell except the integer shape (withheld whatever a window reading says). */
export const scoredCells = (cells: readonly SegmentCell[]): SegmentCell[] => cells.filter(c => !integerShape(c.segment))

export function evaluateUpperWindow(cells: readonly SegmentCell[], candidates: readonly UpperWindowCandidate[] = upperWindowCandidates()): UpperWindowResult[] {
  const scored = scoredCells(cells)
  const out: UpperWindowResult[] = []
  for (const [scope, inScope] of Object.entries(UPPER_WINDOW_SCOPES)) {
    const scoped = scored.filter(c => inScope(c.segment))
    const outside = scored.filter(c => !inScope(c.segment) && hex(truncationOf(c.segment)) !== c.excel).map(keyOf)
    for (const cand of candidates) {
      const mismatches = [...outside]
      for (const c of scoped) { const got = cand.evaluate(c.segment); if (got === undefined || hex(got) !== c.excel) mismatches.push(keyOf(c)) }
      out.push({ id: cand.id, family: cand.family, scope, inScope: scoped.length, mismatches })
    }
  }
  return out
}

export interface WindowCensusRow {
  caseId: string
  ref: string
  colours: 2 | 3
  segment: 0 | 1
  bounds: string
  value: number
  lo: number
  hi: number
  fraction: number
  /** Per RGB channel: d and the double product f·d with its class. */
  channels: Array<{ d: number; product: number; cls: ProductClass }>
  excel: string
  guarded: string
  rawFloor: string
  /** Excel against the two readings: `guarded`, `raw-floor`, `both` (they coincide) or `neither`. */
  outcome: 'guarded' | 'raw-floor' | 'both' | 'neither'
}

/** Every captured interior cell with some channel inside the window but no IEEE-integer channel, at any segment or scale. */
export function windowCensus(cells: readonly SegmentCell[]): WindowCensusRow[] {
  return cells.filter(c => interior(c.segment) && !channelClasses(c.segment).includes('integer') && channelClasses(c.segment).some(x => x === 'below' || x === 'above')).map(c => {
    const s = c.segment
    const a = channelsOf(s.c0), b = channelsOf(s.c1)
    const guarded = hex(truncationOf(s)), rawFloor = hex(READING_CANDIDATES[1]!.evaluate(s)!)
    const outcome = c.excel === guarded && c.excel === rawFloor ? 'both' : c.excel === guarded ? 'guarded' : c.excel === rawFloor ? 'raw-floor' : 'neither'
    return {
      caseId: c.caseId, ref: c.ref, colours: s.colours, segment: s.segment, bounds: `${s.loKind}/${s.hiKind}`, value: s.value, lo: s.lo, hi: s.hi, fraction: s.fraction,
      channels: [0, 1, 2].map(i => { const d = Math.abs(b[i]! - a[i]!); return { d, product: s.fraction * d, cls: productClass(s.fraction, d) } }),
      excel: c.excel, guarded, rawFloor, outcome,
    }
  })
}

export async function upperWindowCells(rootDir: string = REPO_ROOT): Promise<SegmentCell[]> {
  const manifest = JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const { cells, faults } = await collectSegmentCells(manifest, rootDir, { include: 'all' })
  if (faults.length > 0) throw new Error(`upper-window: harness faults ${faults.join('; ')}`)
  return cells
}

// ---------------------------------------------------------------------------------------------------------------------
// The case's predictions (docs/validation/xlsx-cf-upper-window-predictions.md, `--write-predictions`).
// ---------------------------------------------------------------------------------------------------------------------

export interface UpperWindowPrediction {
  ref: string
  configuration: string
  block: string
  bounds: string
  value: string
  /** The segment: L (lower or only) or U (upper). */
  segment: 'L' | 'U'
  /** f·d per RGB channel (d = 0 omitted), with production's double f. */
  products: string
  /** The independent model's class: integer, window or none. */
  cls: 'integer' | 'window' | 'none'
  readings: Record<string, string>
  /** Production today (nothing registered): `painted` or the withhold. */
  production: string
  /** Production with a window reading registered: the reading's colour at a statistical window cell, else as today. */
  registered: string
  /** The prediction tokens per reading (twin or literal). */
  tokens: Record<string, string>
}

/** Every designated cell's predictions under each reading, with production today and under a registration. */
export function upperWindowPredictions(): UpperWindowPrediction[] {
  const { configurations } = s2UpperWindowLayout()
  const tokenOf = (reading: string, ref: string): string => configurations.find(cfg => cfg.refs.includes(ref))!.hypotheses[reading]![ref]!
  return s2UpperWindowCells().filter(c => c.configuration !== undefined).map((c: S2UpperWindowCell): UpperWindowPrediction => {
    const f = (c.value - c.lo) / (c.hi - c.lo)
    const ds = [2, 4, 6].map(i => Math.abs(parseInt(c.to.slice(i, i + 2), 16) - parseInt(c.from.slice(i, i + 2), 16))).filter(d => d !== 0)
    const cls = s2UpperWindowClass(c)
    const literal = c.block.cfvo.length === 3 && c.block.cfvo.every(x => x.type === 'num')
    const numSegment = literal && c.value > c.lo && c.value < c.hi && !Number.isInteger(f * 1048576)
    const production = numSegment ? 'withheld (num-segment)' : cls !== 'none' ? 'withheld (upper whole product)' : 'painted'
    return {
      ref: c.ref, configuration: c.configuration!, block: c.block.id, bounds: c.block.cfvo.map(x => x.type + (x.val !== undefined ? ` ${x.val}` : '')).join(' / '), value: c.text, segment: c.segment === 1 ? 'U' : 'L',
      products: [...new Set(ds)].map(d => `${d}: ${f * d}`).join('; '), cls,
      readings: Object.fromEntries(S2_UPPER_WINDOW_READINGS.map(r => [r, s2UpperWindowPrediction(c, r)])),
      production, registered: production === 'withheld (upper whole product)' && cls === 'window' ? 'reading' : production,
      tokens: Object.fromEntries(S2_UPPER_WINDOW_READINGS.map(r => [r, tokenOf(r, c.ref)])),
    }
  })
}

function writePredictions(): void {
  const rows = upperWindowPredictions()
  const lines = [
    `# ${S2_UPPER_WINDOW_ID}: predictions at every designated cell`,
    '',
    'Generated by `npx tsx scripts/cf-refs/upper-window-candidates.ts --write-predictions` from the catalogue. Two readings are pre-registered for `scale:upper-window-statistical` (the window class of a 3-colour upper segment not bounded by two literal num cfvo):',
    '- **guarded**: floor(f·d + 1e-11), production\'s interpolation;',
    '- **raw-floor**: floor(f·d).',
    '',
    'Both predict guarded truncation outside the window class. Class: `window` (within 1e-11 of a whole k, 0 < k < d, no IEEE integer), `integer` (an IEEE-integer product) or `none` (and every lower-segment or 2-colour cell). **Decisive** marks the cells where the two readings differ. Production is today\'s default; "registered" is production with a window reading registered (`reading`: the registered reading\'s colour).',
    '',
    '| Cell | Config | Block | Bounds | Value | Seg | f·d (double) | Class | guarded | raw-floor | Decisive | Production today | Registered |',
    '|---|---|---|---|---|---|---|---|---|---|---|---|---|',
    ...rows.map(r => `| ${r.ref} | ${r.configuration} | ${r.block} | ${r.bounds} | ${r.value} | ${r.segment} | ${r.products} | ${r.cls} | ${r.readings.guarded}${r.tokens.guarded!.startsWith('twin:') ? ` (${r.tokens.guarded})` : ''} | ${r.readings['raw-floor'] === r.readings.guarded ? '=' : `${r.readings['raw-floor']}${r.tokens['raw-floor']!.startsWith('twin:') ? ` (${r.tokens['raw-floor']})` : ''}`} | ${r.readings['raw-floor'] === r.readings.guarded ? '' : 'yes'} | ${r.production} | ${r.registered} |`),
    '',
  ]
  writeFileSync(resolve(REPO_ROOT, 'docs/validation/xlsx-cf-upper-window-predictions.md'), lines.join('\n'))
}

async function main(): Promise<void> {
  if (process.argv.includes('--write-predictions')) { writePredictions(); return }
  const cells = await upperWindowCells()
  const short = (k: string): string => k.replace(/^stage\d-(scale-)?/, '')
  const census = windowCensus(cells)
  console.log(`## Census\n\n${cells.length} captured colour-scale cells; ${cells.filter(c => integerShape(c.segment)).length} in the integer shape (not scored); ${scoredCells(cells).length} scored.\n`)
  console.log('| Cell | Scale | Bounds | Value | lo | hi | f | d: f·d (class) per channel | Excel | Guarded | Raw floor | Outcome |\n|---|---|---|---|---|---|---|---|---|---|---|---|')
  for (const r of census) console.log(`| ${short(r.caseId)}:${r.ref} | ${r.colours}-colour seg ${r.segment} | ${r.bounds} | ${r.value} | ${r.lo} | ${r.hi} | ${r.fraction} | ${r.channels.map(x => (x.d === 0 ? '0' : `${x.d}: ${x.product} (${x.cls})`)).join('; ')} | ${r.excel} | ${r.guarded} | ${r.rawFloor} | ${r.outcome} |`)
  const results = evaluateUpperWindow(cells)
  const families = ['origin', 'f32', 'precision', 'rational', 'form', 'threshold', 'reading'] as const
  console.log('\n## Survivors\n\n| Family | Scope | Candidates | In-scope cells | Survivors | Fewest mismatches | Best (first 3) and their mismatches |\n|---|---|---|---|---|---|---|')
  for (const fam of families) for (const scope of Object.keys(UPPER_WINDOW_SCOPES)) {
    const rs = results.filter(r => r.family === fam && r.scope === scope)
    const best = Math.min(...rs.map(r => r.mismatches.length))
    const top = rs.filter(r => r.mismatches.length === best)
    console.log(`| ${fam} | ${scope} | ${rs.length} | ${rs[0]!.inScope} | ${rs.filter(r => r.mismatches.length === 0).length} | ${best} (${top.length}) | ${top.slice(0, 3).map(r => `\`${r.id}\``).join(', ')}: ${top[0]!.mismatches.slice(0, 6).map(short).join(', ')}${top[0]!.mismatches.length > 6 ? ', …' : ''} |`)
  }
  const near = results.filter(r => r.scope === 'window' && r.mismatches.length === 1)
  const bySet = new Map<string, UpperWindowResult[]>()
  for (const r of near) bySet.set(r.mismatches[0]!, [...(bySet.get(r.mismatches[0]!) ?? []), r])
  console.log(`\n## Near misses at the window scope (1 mismatch): ${near.length} candidates\n\n| Mismatch | Candidates | Families | Examples |\n|---|---|---|---|`)
  for (const [k, rs] of bySet) console.log(`| ${short(k)} | ${rs.length} | ${[...new Set(rs.map(r => r.family))].join(', ')} | ${rs.slice(0, 3).map(r => `\`${r.id}\``).join(', ')} |`)
  console.log(`\n## The readings\n`)
  for (const r of results.filter(x => x.family === 'reading')) console.log(`- \`${r.id}\` at ${r.scope}: ${r.mismatches.length} mismatches${r.mismatches.length ? ` (${r.mismatches.map(short).join(', ')})` : ''}`)
  console.log(`\nTotal: ${results.length / Object.keys(UPPER_WINDOW_SCOPES).length} candidates x ${Object.keys(UPPER_WINDOW_SCOPES).length} scopes; survivors at any scope: ${results.filter(r => r.mismatches.length === 0).length}.`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
