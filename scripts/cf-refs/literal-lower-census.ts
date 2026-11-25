/**
 * The literal-lower follow-up (brief cf-literal-lower, Phase 1 task 1): the census of every captured interior cell of a
 * 3-colour LOWER segment bounded by two literal `num` cfvo (after the midpoint-above-max re-sort), Stage 2 and the Stage 5
 * painted colour-scale cells (`collectSegmentCells`, `{ include: 'all' }`). Each is compared with production's lower-segment
 * arithmetic as now registered: in the guard-boundary shape (`guardBoundaryProduct`) the registered reading of
 * scale:guard-boundary-raw-floor (`raw-floor`, painted by `upperWindowReadingColor`), elsewhere `interpolateScaleColor`'s
 * truncation. Production's own functions are called, so the census checks what production would paint once the num-segment
 * withhold no longer applies to segment 0. Every such cell must equal Excel's DisplayFormat fill, else the brief stops
 * (BLOCKED).
 *
 * Each row is classified by:
 * - span: the segment's value span hi - lo, with lo (non-zero, negative and fractional bounds are visible);
 * - order: the directions of its channels with d != 0 (`asc`, `desc`, `both`) and its palette `c0>c1`;
 * - kind: `p-line` (every channel with d != 0 has the same |d|) or `distinct` (spans differ per channel);
 * - class: the worst channel product p = f·|d| (`below` / `above` / `integer` / `clear`, as raw-floor-census.ts);
 * - f20: whether f·2^20 is an integer (the num-segment withhold applies only where it is not).
 *
 * Run: `npx tsx scripts/cf-refs/literal-lower-census.ts` (every capture, stage2-scale-literal-lower included: recipe §12.7);
 * `--pre-capture` leaves that capture out (recipe §12.1). Both are pinned by tests/xlsx-cf-scale-literal-lower-census.test.ts.
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { cfGuardBoundaryReading, CF_GUARD_BOUNDARY_CAP, guardBoundaryProduct, interpolateScaleColor, upperWindowReadingColor } from '../../src/xlsx/cf/colorScale'
import type { CfManifest } from './cases'
import { collectSegmentCells, type SegmentCell, type SegmentCollection } from './num-segment-candidates'
import { productClass, type CensusClass } from './raw-floor-census'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const RANK: Record<CensusClass, number> = { below: 3, above: 2, integer: 1, clear: 0 }

export interface LiteralLowerRow {
  caseId: string
  ref: string
  lo: number
  hi: number
  span: number
  value: number
  fraction: number
  palette: string
  order: 'asc' | 'desc' | 'both'
  kind: 'p-line' | 'distinct'
  cls: CensusClass
  /** f·2^20 is an integer (the num-segment withhold does not apply; production paints the cell today). */
  f20: boolean
  excel: string
  /** Production's lower-segment arithmetic as registered (raw floor at guard-boundary products, truncation elsewhere). */
  lower: string
  status: SegmentCell['status']
}

/** Production's lower-segment colour for a segment, as registered today (sRGB, truncate). */
export function productionLowerColour(c0: number, c1: number, value: number, lo: number, hi: number, fraction: number): string {
  if (guardBoundaryProduct(c0, c1, value, lo, hi, fraction)) {
    const r = cfGuardBoundaryReading(CF_GUARD_BOUNDARY_CAP)
    const argb = r !== undefined ? upperWindowReadingColor(r, c0, c1, fraction) : undefined
    if (argb === undefined) throw new Error('guard-boundary reading unregistered or non-finite')
    return hex(argb)
  }
  return hex(interpolateScaleColor(c0, c1, fraction))
}

/** The interior literal-num lower-segment rows of a collection. */
export function literalLowerCensus(collection: SegmentCollection): { rows: LiteralLowerRow[]; endpoints: number } {
  const rows: LiteralLowerRow[] = []
  let endpoints = 0
  for (const x of collection.cells) {
    const s = x.segment
    if (s.colours !== 3 || s.segment !== 0 || s.loKind !== 'num' || s.hiKind !== 'num') continue
    if (!(s.value > s.lo && s.value < s.hi)) { endpoints++; continue }
    let cls: CensusClass = 'clear'
    const dirs = new Set<'asc' | 'desc'>(), spans = new Set<number>()
    for (const shift of [24, 16, 8, 0]) {
      const a = (s.c0 >>> shift) & 0xff, b = (s.c1 >>> shift) & 0xff
      if (a === b) continue
      dirs.add(b > a ? 'asc' : 'desc')
      spans.add(Math.abs(b - a))
      const c = productClass(s.fraction * Math.abs(b - a))
      if (RANK[c] > RANK[cls]) cls = c
    }
    rows.push({
      caseId: x.caseId, ref: x.ref, lo: s.lo, hi: s.hi, span: s.hi - s.lo, value: s.value, fraction: s.fraction,
      palette: `${hex(s.c0)}>${hex(s.c1)}`, order: dirs.size === 2 ? 'both' : dirs.has('desc') ? 'desc' : 'asc',
      kind: spans.size <= 1 ? 'p-line' : 'distinct', cls, f20: Number.isInteger(s.fraction * 1048576), excel: x.excel,
      lower: productionLowerColour(s.c0, s.c1, s.value, s.lo, s.hi, s.fraction), status: x.status,
    })
  }
  return { rows, endpoints }
}

/**
 * The capture that followed this census (stage2-scale-literal-lower, a3f85309): `{ include: 'pre-capture' }` leaves it out,
 * reproducing the numbers of recipe §12.1 as they were when the case was designed; the default includes it (recipe §12.7).
 */
export const LITERAL_LOWER_POST_CAPTURE_CASES: readonly string[] = ['stage2-scale-literal-lower']
/** Captures after this census's own follow-up, left out in every mode: the census is recorded as of a3f85309 (recipe §12.7). */
export const LITERAL_LOWER_LATER_CASES: readonly string[] = ['stage2-scale-upper-segment', 'stage2-scale-formula-bounds', 'stage2-scale-percent-order']

export async function literalLowerCensusOf(manifest?: CfManifest, rootDir: string = REPO_ROOT, opts: { include?: 'pre-capture' | 'all' } = {}): Promise<{ collection: SegmentCollection } & ReturnType<typeof literalLowerCensus>> {
  const m = manifest ?? (JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest)
  const all = await collectSegmentCells(m, rootDir, { include: 'all' })
  const keep = (x: { caseId: string }): boolean => !LITERAL_LOWER_LATER_CASES.includes(x.caseId) && ((opts.include ?? 'all') === 'all' || !LITERAL_LOWER_POST_CAPTURE_CASES.includes(x.caseId))
  const collection: SegmentCollection = { cells: all.cells.filter(keep), excluded: all.excluded.filter(keep), faults: all.faults }
  return { collection, ...literalLowerCensus(collection) }
}

/** Counts by a key: cells, and how many production's lower-segment arithmetic paints at Excel's fill. */
export function countBy(rows: readonly LiteralLowerRow[], key: (r: LiteralLowerRow) => string): Array<{ key: string; cells: number; exact: number }> {
  const m = new Map<string, { cells: number; exact: number }>()
  for (const r of rows) {
    const k = key(r)
    const e = m.get(k) ?? { cells: 0, exact: 0 }
    e.cells++
    if (r.lower === r.excel) e.exact++
    m.set(k, e)
  }
  return [...m].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)).map(([k, v]) => ({ key: k, ...v }))
}

async function main(): Promise<void> {
  const { collection, rows, endpoints } = await literalLowerCensusOf(undefined, REPO_ROOT, { include: process.argv.includes('--pre-capture') ? 'pre-capture' : 'all' })
  const withheld = rows.filter(r => !r.f20)
  console.log(`segment cells ${collection.cells.length}; harness faults ${collection.faults.length}; literal-num lower-segment interior ${rows.length} (f·2^20 not an integer ${withheld.length}); endpoints or clamped ${endpoints}`)
  const table = (title: string, key: (r: LiteralLowerRow) => string, rs: readonly LiteralLowerRow[]): void => {
    console.log(title)
    for (const t of countBy(rs, key)) console.log(`  ${t.key.padEnd(48)} ${String(t.cells).padStart(4)}  exact ${t.exact}`)
  }
  for (const [label, rs] of [['all interior', rows], ['f·2^20 not an integer (withheld today)', withheld]] as const) {
    console.log(`== ${label}`)
    table('by case', r => r.caseId, rs)
    table('by span (lo..hi)', r => `${r.span} (${r.lo}..${r.hi})`, rs)
    table('by colour order and palette', r => `${r.order} ${r.palette}`, rs)
    table('by palette kind', r => r.kind, rs)
    table('by product class', r => r.cls, rs)
    table('by kind × class', r => `${r.kind} ${r.cls}`, rs)
  }
  const misses = rows.filter(r => r.lower !== r.excel)
  console.log(`misses: ${misses.length}`)
  for (const r of misses) console.log(`  ${r.caseId}:${r.ref} v=${r.value} [${r.lo}, ${r.hi}] f=${r.fraction} ${r.palette} ${r.cls} Excel ${r.excel} lower ${r.lower}`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
