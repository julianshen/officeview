/**
 * The formula-bounds follow-up (brief cf-formula-bounds, Phase 1 task 1): the census of every captured colour-scale cell
 * whose segment is bounded by at least one formula-resolved cfvo (`num:formula`, `percent:formula`, `percentile:formula`).
 *
 * Cells: `collectSegmentCells` (scripts/cf-refs/num-segment-candidates.ts, `{ include: 'all' }`): every measured case with a
 * colour scale, each painted cell mapped to the segment production evaluated for it. A cell is in the census when either
 * bound of its segment is formula-resolved.
 *
 * Per cell: the shape production's segment rules put it in (as if its bounds were literal: `guard-boundary`, `num-segment`,
 * `upper-integer`, `upper-window`, `upper-near-top`, `clear`, `endpoint`); the bound kinds; each formula bound's resolved
 * double (the formula result: the threshold itself for `num:formula`, the p of a `percent:formula` / `percentile:formula`
 * recomputed from the sheet) in IEEE hex, its 15-significant-digit rounding (`Number(x.toPrecision(15))`) and whether it
 * differs; Excel's DisplayFormat fill; production's paint (undefined: withheld); and each reading's prediction:
 * - `as-literal`: the formula bound's double substituted as the same literal (production's segment rules, every registered
 *   reading, with the bound kinds stripped of `:formula`);
 * - `fifteen-digit`: the same with every formula result first rounded to 15 significant digits (thresholds recomputed).
 *
 * Nothing here is registered or reachable from production.
 *
 * Run: `npx tsx scripts/cf-refs/formula-bounds-census.ts` (prints the table of docs/validation/xlsx-cf-recipes-stage2.md §14.1);
 * `--all` includes the capture of stage2-scale-formula-bounds (recipe §14.7).
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { guardBoundaryProduct, numSegmentUnmeasured, safeFraction, scaleColorAt, upperSegmentNearTop, upperSegmentProductClass, type CfScaleSegment, type CfScaleSegmentBoundKind } from '../../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../../src/xlsx/cf/model'
import { caseParseOptionsOf } from './case-clock'
import type { CfManifest } from './cases'
import { collectSegmentCells, type SegmentCollection } from './num-segment-candidates'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const f64 = new Float64Array(1), u64 = new BigUint64Array(f64.buffer)
/** The IEEE bits of a double, 16 hex digits. */
export const bitsOf = (x: number): string => { f64[0] = x; return u64[0]!.toString(16).padStart(16, '0') }
/** The `fifteen-digit` rounding of a formula result: the double nearest its 15-significant-digit decimal. */
export const fifteenDigit = (x: number): number => Number(x.toPrecision(15))

export type FormulaBoundsShape = 'endpoint' | 'guard-boundary' | 'num-segment' | 'upper-integer' | 'upper-window' | 'upper-near-top' | 'clear'

export interface FormulaBoundsRow {
  caseId: string
  ref: string
  value: number
  colours: 2 | 3
  segment: 0 | 1
  bounds: string
  shape: FormulaBoundsShape
  /** Each formula-resolved bound of the segment: which end, its kind, the formula result, its hex, its 15-digit rounding. */
  formula: Array<{ end: 'lo' | 'hi'; kind: CfScaleSegmentBoundKind; result: number | undefined; resultHex: string; rounded: number | undefined; roundedHex: string; differs: boolean | undefined }>
  excel: string
  production: string | undefined
  asLiteral: string | undefined
  fifteenDigit: string | undefined
  /** Whether the two readings predict different fills (they can tell each other apart here). */
  discriminating: boolean
}

const stripKind = (kind: CfScaleSegmentBoundKind, value: number, keep = false): CfScalePoint['bound'] =>
  kind === 'min' || kind === 'max' ? { type: kind } : { type: kind.split(':')[0] as 'num' | 'percent' | 'percentile', value, ...(keep && kind.endsWith(':formula') ? { resolved: 'formula' as const } : {}) }

/**
 * The colour production's segment rules give `seg` with literal bounds (every registered reading with its evidence): a
 * synthetic ascending scale holding the segment with the segment's colours, its third threshold (3-colour) one of the scale's.
 * Undefined: withheld. `keepFormula`: the bounds keep their formula provenance (production's own paint of the segment).
 */
export function literalSegmentColour(seg: CfScaleSegment, lo: number = seg.lo, hi: number = seg.hi, keepFormula = false): string | undefined {
  const loB = stripKind(seg.loKind, lo, keepFormula), hiB = stripKind(seg.hiKind, hi, keepFormula)
  let points: CfScalePoint[], thresholds: number[]
  if (seg.colours === 2) { points = [{ bound: loB, argb: seg.c0 }, { bound: hiB, argb: seg.c1 }]; thresholds = [lo, hi] }
  else if (seg.segment === 0) { points = [{ bound: loB, argb: seg.c0 }, { bound: hiB, argb: seg.c1 }, { bound: { type: 'max' }, argb: seg.c1 }]; thresholds = [lo, hi, Math.max(seg.thresholds[2]!, hi + 1)] }
  else { points = [{ bound: { type: 'min' }, argb: seg.c0 }, { bound: loB, argb: seg.c0 }, { bound: hiB, argb: seg.c1 }]; thresholds = [Math.min(seg.thresholds[0]!, lo - 1), lo, hi] }
  const out = scaleColorAt(seg.value, points, thresholds)
  return out.kind === 'ready' ? hex(out.value) : undefined
}

/** The shape of a segment read with literal bounds (production's precedence: near-top, num-segment, upper classes, guard boundary). */
export function literalShape(seg: CfScaleSegment, lo: number = seg.lo, hi: number = seg.hi): FormulaBoundsShape {
  const fraction = safeFraction(seg.value, lo, hi) ?? seg.fraction
  if (!(seg.value > lo && seg.value < hi)) return 'endpoint'
  const p = (kind: CfScaleSegmentBoundKind, v: number): CfScalePoint => ({ bound: stripKind(kind, v), argb: 0 })
  if (seg.colours === 3 && seg.segment === 1 && upperSegmentNearTop(seg.c0, seg.c1, seg.value, lo, hi, fraction)) return 'upper-near-top'
  if (seg.colours === 3 && numSegmentUnmeasured(p(seg.loKind, lo), p(seg.hiKind, hi), seg.value, lo, hi, fraction)) return 'num-segment'
  if (seg.colours === 3 && seg.segment === 1) {
    const cls = upperSegmentProductClass(seg.c0, seg.c1, seg.value, lo, hi, fraction)
    if (cls) return cls === 'integer' ? 'upper-integer' : 'upper-window'
    return 'clear'
  }
  return guardBoundaryProduct(seg.c0, seg.c1, seg.value, lo, hi, fraction) ? 'guard-boundary' : 'clear'
}

type SheetNumbers = ReadonlyMap<string, number>
/** A `percent` / `percentile` formula result recomputed from the sheet: an absolute one-cell reference holding a number. */
function percentResult(formula: string | undefined, cells: SheetNumbers): number | undefined {
  const m = /^=?\$?([A-Z]+)\$?(\d+)$/.exec((formula ?? '').trim())
  return m ? cells.get(`${m[1]}${m[2]}`) : undefined
}

/**
 * The capture that followed this census (stage2-scale-formula-bounds, d2660d46): `{ include: 'pre-capture' }` (the default,
 * the census as run before the capture, recipe §14.1) leaves it out; `'all'` includes it (recipe §14.7).
 */
export const FORMULA_BOUNDS_POST_CAPTURE_CASES: readonly string[] = ['stage2-scale-formula-bounds']

export async function formulaBoundsCensusOf(manifest?: CfManifest, rootDir: string = REPO_ROOT, opts: { include?: 'pre-capture' | 'all' } = {}): Promise<{ collection: SegmentCollection; rows: FormulaBoundsRow[] }> {
  const m = manifest ?? (JSON.parse(readFileSync(resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest)
  const everything = await collectSegmentCells(m, rootDir, { include: 'all' })
  const keep = (x: { caseId: string }): boolean => (opts.include ?? 'pre-capture') === 'all' || !FORMULA_BOUNDS_POST_CAPTURE_CASES.includes(x.caseId)
  const collection: SegmentCollection = { cells: everything.cells.filter(keep), excluded: everything.excluded.filter(keep), faults: everything.faults }
  const rows: FormulaBoundsRow[] = []
  const sheets = new Map<string, SheetNumbers>()
  const numbersOf = async (caseId: string): Promise<SheetNumbers> => {
    if (sheets.has(caseId)) return sheets.get(caseId)!
    const c = m.cases.find(x => x.id === caseId)!
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(rootDir, c.saved)))), c.stage === 5 ? caseParseOptionsOf(c) : {})
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const out = new Map<string, number>()
    for (const r of sheet.rows) for (const cl of r.cells) if (typeof cl.value === 'number') out.set(`${colName(cl.col)}${r.index + 1}`, cl.value)
    sheets.set(caseId, out)
    return out
  }
  for (const x of collection.cells) {
    const s = x.segment
    if (!s.loKind.endsWith(':formula') && !s.hiKind.endsWith(':formula')) continue
    const c = m.cases.find(y => y.id === x.caseId)!
    const marker = c.markers.find(k => k.ruleType === 'colorScale' && k.scale && sqrefCovers(k.sqref, x.ref))
    const numbers = await numbersOf(x.caseId)
    const scope = [...numbers].filter(([k]) => marker && sqrefCovers(marker.sqref, k)).map(([, v]) => v).sort((a, b) => a - b)
    const formula: FormulaBoundsRow['formula'] = []
    let lo15 = s.lo, hi15 = s.hi, resolvable = true
    for (const end of ['lo', 'hi'] as const) {
      const kind = end === 'lo' ? s.loKind : s.hiKind
      if (!kind.endsWith(':formula')) continue
      const t = end === 'lo' ? s.lo : s.hi
      let result: number | undefined
      if (kind === 'num:formula') result = t
      else {
        // The formula cfvo of this kind in the marker (the only one of its kind in every captured rule).
        const cfvo = marker?.scale?.cfvo.filter(v => v.type === kind.split(':')[0] && (v as { f?: string }).f !== undefined)
        result = cfvo?.length === 1 ? percentResult((cfvo[0] as { f?: string }).f, numbers) : undefined
      }
      const rounded = result === undefined ? undefined : fifteenDigit(result)
      formula.push({ end, kind, result, resultHex: result === undefined ? '?' : bitsOf(result), rounded, roundedHex: rounded === undefined ? '?' : bitsOf(rounded), differs: result === undefined ? undefined : rounded !== result })
      if (result === undefined || rounded === undefined) { resolvable = false; continue }
      if (rounded === result) continue
      // The threshold from the rounded result: num as itself; percent min + p/100·(max − min); percentile PERCENTILE.INC.
      const min = scope[0]!, max = scope.at(-1)!
      const t15 = kind === 'num:formula' ? rounded : kind === 'percent:formula' ? min + (rounded / 100) * (max - min) : percentileInc(scope, rounded)
      if (end === 'lo') lo15 = t15; else hi15 = t15
    }
    const asLiteral = literalSegmentColour(s)
    const moved = lo15 !== s.lo || hi15 !== s.hi
    // A rounded threshold that moves the value out of the segment has no prediction here (none of the census does).
    const fifteen = !resolvable ? undefined : !moved ? asLiteral : s.value >= lo15 && s.value <= hi15 && lo15 < hi15 ? literalSegmentColour(s, lo15, hi15) : undefined
    rows.push({
      caseId: x.caseId, ref: x.ref, value: s.value, colours: s.colours, segment: s.segment, bounds: `${s.loKind}/${s.hiKind}`, shape: literalShape(s),
      formula, excel: x.excel, production: literalSegmentColour(s, s.lo, s.hi, true), asLiteral, fifteenDigit: fifteen,
      discriminating: resolvable && fifteen !== asLiteral,
    })
  }
  return { collection, rows }
}

function percentileInc(sorted: readonly number[], p: number): number {
  const rank = (p / 100) * (sorted.length - 1)
  const lo = Math.floor(rank)
  return lo + 1 < sorted.length ? sorted[lo]! + (rank - lo) * (sorted[lo + 1]! - sorted[lo]!) : sorted[lo]!
}

function colName(col: number): string {
  let s = ''
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}
const rc = (ref: string): { row: number; col: number } => {
  const mm = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(mm[2]) - 1, col: [...mm[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}
function sqrefCovers(sqref: string, ref: string): boolean {
  const { row, col } = rc(ref)
  return sqref.split(/\s+/).some(part => {
    const [a, b = a] = part.split(':')
    const p = rc(a!.replace(/\$/g, '')), q = rc(b!.replace(/\$/g, ''))
    return row >= Math.min(p.row, q.row) && row <= Math.max(p.row, q.row) && col >= Math.min(p.col, q.col) && col <= Math.max(p.col, q.col)
  })
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { collection, rows } = await formulaBoundsCensusOf(undefined, REPO_ROOT, { include: process.argv.includes('--all') ? 'all' : 'pre-capture' })
  console.log(`segment cells ${collection.cells.length}, faults ${collection.faults.length}; formula-bounded ${rows.length}`)
  const count = (key: (r: FormulaBoundsRow) => string): string => {
    const t = new Map<string, number>()
    for (const r of rows) t.set(key(r), (t.get(key(r)) ?? 0) + 1)
    return [...t].sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, n]) => `${k} ${n}`).join('; ')
  }
  console.log('by shape:', count(r => r.shape))
  console.log('by bounds:', count(r => `${r.colours}c seg${r.segment} ${r.bounds}`))
  console.log('by case:', count(r => r.caseId))
  const doubles = new Map<string, { result: number; rounded: number; differs: boolean }>()
  for (const r of rows) for (const b of r.formula) if (b.result !== undefined) doubles.set(`${b.kind} ${b.resultHex}`, { result: b.result, rounded: b.rounded!, differs: b.differs! })
  console.log('distinct formula results:')
  for (const [k, v] of [...doubles].sort()) console.log(`  ${k} = ${v.result}  15-digit ${v.rounded} (${bitsOf(v.rounded)}) ${v.differs ? 'DIFFERS' : 'same'}`)
  const painted = rows.filter(r => r.production !== undefined)
  const disagreeLiteral = painted.filter(r => r.asLiteral !== r.excel)
  console.log(`painted ${painted.length}; painted disagreeing with as-literal ${disagreeLiteral.length}${disagreeLiteral.length ? ': ' + disagreeLiteral.map(r => `${r.caseId}:${r.ref}`).join(', ') : ''}`)
  const withheld = rows.filter(r => r.production === undefined)
  console.log(`withheld ${withheld.length}: ${withheld.map(r => `${r.caseId}:${r.ref} ${r.shape} excel ${r.excel} as-literal ${r.asLiteral ?? '-'} fifteen-digit ${r.fifteenDigit ?? '-'}`).join('; ')}`)
  console.log(`as-literal = Excel at ${rows.filter(r => r.asLiteral === r.excel).length}/${rows.length}; fifteen-digit = Excel at ${rows.filter(r => r.fifteenDigit === r.excel).length}/${rows.length}; discriminating ${rows.filter(r => r.discriminating).length}`)
  if (process.argv.includes('--rows')) for (const r of rows) console.log([r.caseId, r.ref, r.value, `${r.colours}c/s${r.segment}`, r.bounds, r.shape, r.formula.map(b => `${b.end}=${b.result}(${b.resultHex}) r15 ${b.roundedHex}${b.differs ? '*' : ''}`).join(' '), `excel ${r.excel}`, `prod ${r.production ?? '-'}`, `lit ${r.asLiteral ?? '-'}`, `15d ${r.fifteenDigit ?? '-'}`].join(' | '))
}
