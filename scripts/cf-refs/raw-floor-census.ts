/**
 * The raw-floor follow-up (brief cf-rawfloor, Phase 1 task 1): the census of hypothesis W1 (docs/validation/xlsx-cf.md
 * §14.z.17) over every captured interior cell of a 2-colour scale or of a 3-colour LOWER segment (after the
 * midpoint-above-max re-sort): Excel draws raw floor, per ARGB channel the start colour moved by floor(f·|d|) toward the
 * end colour, with f production's `safeFraction` double and |d| the channel's span (no 1e-11 guard).
 *
 * The cells are those of `collectSegmentCells` (scripts/cf-refs/num-segment-candidates.ts, `{ include: 'all' }`): every
 * measured case with a colour scale, Stage 2 and the Stage 5 painted colour-scale cells, each mapped to the segment
 * production evaluated for it. A cell is interior when lo < v < hi; an endpoint or a clamped value is counted apart. Each
 * cell is classified by:
 * - scale: `2c` or `3c-lower`;
 * - bounds: `literal` (two literal num cfvo), `statistical` (min, max, percent, percentile), `formula` (either bound a
 *   formula-resolved cfvo), `mixed` (one literal num, one statistical);
 * - order: the directions of its channels with d != 0 (`asc`, `desc` or `both`);
 * - class: the worst channel product p = f·|d|: `below` (within 1e-11 strictly below a whole number), `above` (within 1e-11
 *   above one), `integer` (an IEEE integer), `clear` (none of these).
 * Raw floor must equal Excel's DisplayFormat fill at every interior cell (W1 exact), else the brief stops (BLOCKED).
 * Guarded truncation (floor(f·|d| + 1e-11), production's `interpolateChannel`) is reported alongside. Nothing here is
 * registered or reachable from production.
 *
 * Run: `npx tsx scripts/cf-refs/raw-floor-census.ts` (about 30 s; every capture, stage2-scale-raw-floor included: recipe
 * §11.7); `--pre-capture` leaves that capture out and prints the tables of docs/validation/xlsx-cf-recipes-stage2.md §11.1.
 * Both are pinned by tests/xlsx-cf-scale-raw-floor-census.test.ts.
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CfManifest } from './cases'
import { collectSegmentCells, type SegmentCell, type SegmentCollection } from './num-segment-candidates'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
/** The guard window, restated (not imported from production): `CF_TRUNCATION_GUARD`. */
const GUARD = 1e-11

export type CensusClass = 'below' | 'above' | 'integer' | 'clear'
export type CensusBounds = 'literal' | 'statistical' | 'formula' | 'mixed'

export interface CensusRow {
  caseId: string
  ref: string
  scale: '2c' | '3c-lower'
  bounds: CensusBounds
  order: 'asc' | 'desc' | 'both'
  cls: CensusClass
  excel: string
  rawFloor: string
  guarded: string
  /** Production's status in the harness (`confirmed` painted Excel's fill; `withheld`). */
  status: SegmentCell['status']
}

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')

/** One channel product's class. */
export function productClass(p: number): CensusClass {
  if (Number.isInteger(p)) return 'integer'
  const k = Math.round(p)
  if (Math.abs(p - k) <= GUARD) return p < k ? 'below' : 'above'
  return 'clear'
}

/** A reading's colour from `c0` toward `c1` at fraction f, per ARGB channel: raw floor (W1) or guarded truncation. */
export function censusColour(reading: 'raw-floor' | 'guarded', c0: number, c1: number, f: number): string {
  let out = 0
  for (const shift of [24, 16, 8, 0]) {
    const a = (c0 >>> shift) & 0xff, b = (c1 >>> shift) & 0xff
    const p = f * Math.abs(b - a)
    const k = reading === 'raw-floor' ? Math.floor(p) : Math.floor(p + GUARD)
    out |= (b >= a ? a + k : a - k) << shift
  }
  return hex(out)
}

const boundKind = (k: string): 'L' | 'S' | 'F' => (k.endsWith(':formula') ? 'F' : k === 'num' ? 'L' : 'S')
function boundsOf(lo: string, hi: string): CensusBounds {
  const a = boundKind(lo), b = boundKind(hi)
  if (a === 'F' || b === 'F') return 'formula'
  if (a === 'L' && b === 'L') return 'literal'
  if (a === 'S' && b === 'S') return 'statistical'
  return 'mixed'
}
const RANK: Record<CensusClass, number> = { below: 3, above: 2, integer: 1, clear: 0 }

/** The interior 2-colour and lower-segment rows of a collection, and the counts of the cells left out. */
export function rawFloorCensus(collection: SegmentCollection): { rows: CensusRow[]; endpoints: number; endpointsRawExact: number; upper: number } {
  const lower = collection.cells.filter(x => x.segment.colours === 2 || x.segment.segment === 0)
  const rows: CensusRow[] = []
  let endpoints = 0, endpointsRawExact = 0
  for (const x of lower) {
    const s = x.segment
    const raw = censusColour('raw-floor', s.c0, s.c1, s.fraction)
    if (!(s.value > s.lo && s.value < s.hi)) { endpoints++; if (raw === x.excel) endpointsRawExact++; continue }
    let cls: CensusClass = 'clear'
    const dirs = new Set<'asc' | 'desc'>()
    for (const shift of [24, 16, 8, 0]) {
      const a = (s.c0 >>> shift) & 0xff, b = (s.c1 >>> shift) & 0xff
      if (a === b) continue
      dirs.add(b > a ? 'asc' : 'desc')
      const c = productClass(s.fraction * Math.abs(b - a))
      if (RANK[c] > RANK[cls]) cls = c
    }
    rows.push({
      caseId: x.caseId, ref: x.ref, scale: s.colours === 2 ? '2c' : '3c-lower', bounds: boundsOf(s.loKind, s.hiKind),
      order: dirs.size === 2 ? 'both' : dirs.has('desc') ? 'desc' : 'asc', cls, excel: x.excel, rawFloor: raw,
      guarded: censusColour('guarded', s.c0, s.c1, s.fraction), status: x.status,
    })
  }
  return { rows, endpoints, endpointsRawExact, upper: collection.cells.length - lower.length }
}

/**
 * The capture that followed this census (stage2-scale-raw-floor, ce3f701d): `{ include: 'pre-capture' }` leaves it out,
 * reproducing the numbers of recipe §11.1 as they were when the case was designed; the default includes it (recipe §11.7).
 */
export const RAW_FLOOR_POST_CAPTURE_CASES: readonly string[] = ['stage2-scale-raw-floor']
/**
 * Captures after this census's own follow-up, left out in every mode: the census is recorded as of ce3f701d (recipe §11.7).
 * stage2-scale-literal-lower (a3f85309) is counted by scripts/cf-refs/literal-lower-census.ts.
 */
export const RAW_FLOOR_LATER_CASES: readonly string[] = ['stage2-scale-literal-lower', 'stage2-scale-upper-segment']

/** Every captured cell of the manifest's measured colour-scale cases, through production's own pipeline. */
export async function rawFloorCensusOf(manifest?: CfManifest, rootDir: string = REPO_ROOT, opts: { include?: 'pre-capture' | 'all' } = {}): Promise<{ collection: SegmentCollection } & ReturnType<typeof rawFloorCensus>> {
  const m = manifest ?? (JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest)
  const all = await collectSegmentCells(m, rootDir, { include: 'all' })
  const keep = (x: { caseId: string }): boolean => !RAW_FLOOR_LATER_CASES.includes(x.caseId) && ((opts.include ?? 'all') === 'all' || !RAW_FLOOR_POST_CAPTURE_CASES.includes(x.caseId))
  const collection: SegmentCollection = { cells: all.cells.filter(keep), excluded: all.excluded.filter(keep), faults: all.faults }
  return { collection, ...rawFloorCensus(collection) }
}

/** Counts by (scale, bounds, order, class): cells, raw-floor exact, guarded exact. */
export function censusTable(rows: readonly CensusRow[]): Array<{ key: string; cells: number; rawFloor: number; guarded: number }> {
  const m = new Map<string, { cells: number; rawFloor: number; guarded: number }>()
  for (const r of rows) {
    const key = `${r.scale} ${r.bounds} ${r.order} ${r.cls}`
    const e = m.get(key) ?? { cells: 0, rawFloor: 0, guarded: 0 }
    e.cells++
    if (r.rawFloor === r.excel) e.rawFloor++
    if (r.guarded === r.excel) e.guarded++
    m.set(key, e)
  }
  return [...m].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)).map(([key, v]) => ({ key, ...v }))
}

/**
 * Per channel (each channel with d != 0 of each interior row's segment), by (scale, direction, class): channels, and how many
 * Excel drew at the raw-floor step and at the guarded step. Both truncation directions are covered when both directions
 * hold below-integer channels.
 */
export function channelTable(collection: SegmentCollection): Array<{ key: string; channels: number; rawFloor: number; guarded: number }> {
  const m = new Map<string, { channels: number; rawFloor: number; guarded: number }>()
  for (const x of collection.cells) {
    const s = x.segment
    if (!(s.colours === 2 || s.segment === 0) || !(s.value > s.lo && s.value < s.hi)) continue
    const excel = Number.parseInt(x.excel, 16)
    for (const shift of [24, 16, 8, 0]) {
      const a = (s.c0 >>> shift) & 0xff, b = (s.c1 >>> shift) & 0xff
      if (a === b) continue
      const p = s.fraction * Math.abs(b - a)
      const key = `${s.colours === 2 ? '2c' : '3c-lower'} ${b > a ? 'asc' : 'desc'} ${productClass(p)}`
      const e = m.get(key) ?? { channels: 0, rawFloor: 0, guarded: 0 }
      const got = (excel >>> shift) & 0xff
      const step = (k: number): number => (b > a ? a + k : a - k)
      e.channels++
      if (step(Math.floor(p)) === got) e.rawFloor++
      if (step(Math.floor(p + GUARD)) === got) e.guarded++
      m.set(key, e)
    }
  }
  return [...m].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)).map(([key, v]) => ({ key, ...v }))
}

async function main(): Promise<void> {
  const { collection, rows, endpoints, endpointsRawExact, upper } = await rawFloorCensusOf(undefined, REPO_ROOT, { include: process.argv.includes('--pre-capture') ? 'pre-capture' : 'all' })
  const misses = rows.filter(r => r.rawFloor !== r.excel)
  console.log(`segment cells ${collection.cells.length}; 2-colour or lower-segment interior ${rows.length}; endpoints or clamped ${endpoints} (raw floor exact ${endpointsRawExact}); upper-segment ${upper}; harness faults ${collection.faults.length}`)
  const gated = collection.excluded.filter(e => e.group === 'b')
  console.log(`not covered (numeric, filled by Excel, withheld by another production gate before interpolation): ${gated.length}, ${gated.filter(e => e.threeColour).length} under a 3-colour scale`)
  for (const t of censusTable(rows)) console.log(`${t.key.padEnd(34)} ${String(t.cells).padStart(4)}  raw floor ${t.rawFloor}  guarded ${t.guarded}`)
  console.log('per channel:')
  for (const t of channelTable(collection)) console.log(`  ${t.key.padEnd(24)} ${String(t.channels).padStart(4)}  raw floor ${t.rawFloor}  guarded ${t.guarded}`)
  for (const cls of ['below', 'above'] as const) console.log(`${cls}: ${rows.filter(r => r.cls === cls).map(r => `${r.caseId.replace(/^stage\d-(scale-)?/, '')}:${r.ref}`).join(' ')}`)
  console.log(`W1 (raw floor) misses: ${misses.length}`)
  for (const r of misses) console.log(`  ${r.caseId}:${r.ref} ${r.scale} ${r.bounds} ${r.cls} Excel ${r.excel} raw ${r.rawFloor} guarded ${r.guarded}`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
