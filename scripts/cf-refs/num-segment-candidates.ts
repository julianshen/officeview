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
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { guardBoundaryProduct, injectCfScaleSegmentProbe, interpolateScaleColor, numSegmentUnmeasured, upperSegmentWholeProduct, type CfScaleSegment } from '../../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../../src/xlsx/cf/model'
import { caseDocumentEvaluationOf, caseParseOptionsOf } from './case-clock'
import { FRACTIONS, numSegmentCandidates, SWEEP_N, type Candidate } from './num-segment-formulas'
import {
  S2_NUM_SEGMENT,
  S2_NUM_SEGMENT_CAP,
  s2NumSegmentLayout,
  s2NumSegmentRules,
  s2NumSegmentValueAt,
  type CfManifest,
  type CfOracleCase,
  type S2NumSegmentRule,
} from './cases'

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
  /**
   * Captured cells under a colour scale that are not in `cells` (fix round 1 R3), each in one group:
   * - `a`, inapplicable: no DisplayFormat fill observation, Excel shows no fill, a non-numeric value, an endpoint or
   *   equal-threshold colour (no interpolation), or a fill painted by another rule over the scale (obscured);
   * - `b`, a numeric cell Excel fills that is dropped only because a production gate blocks it (overlap, hidden row,
   *   stub, non-numeric bound, composition); `threeColour` says whether any 3-colour scale covers it.
   */
  excluded: Array<{ caseId: string; ref: string; reason: string; group: 'a' | 'b'; threeColour?: boolean }>
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
/**
 * The captures the pre-registration search ran on: every measured case but stage2-scale-num-segment itself (its own
 * confirmation case, captured after the readings were fixed) and the later stage2-scale-upper-whole (54fc3388),
 * stage2-scale-upper-window (ed0fccb6), stage2-scale-raw-floor (ce3f701d) and stage2-scale-literal-lower (a3f85309). Pass
 * `{ include: 'all' }` to add them.
 */
export const SEARCH_EXCLUDED_CASES: readonly string[] = ['stage2-scale-num-segment', 'stage2-scale-upper-whole', 'stage2-scale-upper-window', 'stage2-scale-raw-floor', 'stage2-scale-literal-lower', 'stage2-scale-upper-segment', 'stage2-scale-formula-bounds', 'stage2-scale-percent-order']

/**
 * Captures no census or search of this file's family was run on: stage2-scale-percent-order (ecf00906) postdates them all, and
 * its registration (§14.z.24) changed production's percent arithmetic, so the censuses, kept as of their captures, leave it
 * out even under `{ include: 'all' }`.
 */
export const SEGMENT_CENSUS_LATER_CASES: readonly string[] = ['stage2-scale-percent-order']

export async function collectSegmentCells(manifest: CfManifest, rootDir: string = REPO_ROOT, opts: { include?: 'search' | 'all' } = {}): Promise<SegmentCollection> {
  const out: SegmentCollection = { cells: [], excluded: [], faults: [] }
  for (const c of manifest.cases) {
    if (c.status !== 'measured' || !c.markers.some(m => m.ruleType === 'colorScale')) continue
    if (SEGMENT_CENSUS_LATER_CASES.includes(c.id)) continue
    if ((opts.include ?? 'search') === 'search' && SEARCH_EXCLUDED_CASES.includes(c.id)) continue
    const fills = excelFills(c, rootDir)
    if (!fills) { out.excluded.push({ caseId: c.id, ref: '*', reason: 'no DisplayFormat fill observation in the capture (bar/icon sidecar)', group: 'a' }); continue }
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
      const { row, col } = rc(ref)
      const value = sheet.rows.find(r => r.index === row)?.cells.find(x => x.col === col)?.value
      const threeColour = c.markers.some(m => m.ruleType === 'colorScale' && m.scale?.cfvo.length === 3 && sqrefCovers(m.sqref, ref))
      const exclude = (reason: string, gated: boolean): void => {
        const inapplicable = !gated || excel === 'FFFFFFFF' || typeof value !== 'number'
        const why = excel === 'FFFFFFFF' && gated ? `${reason}; Excel shows no fill` : typeof value !== 'number' && gated ? `${reason}; non-numeric value` : reason
        out.excluded.push({ caseId: c.id, ref, reason: why, group: inapplicable ? 'a' : 'b', threeColour })
      }
      if (!seg) {
        if (scaleCovered) {
          if (production.get(ref) === undefined) exclude('withheld before interpolation (another gate)', true)
          else exclude('painted without a segment evaluation (equal thresholds return an endpoint colour, or another rule)', false)
        }
        continue
      }
      const prod = production.get(ref)
      const truncate = hex(interpolateScaleColor(seg.c0, seg.c1, seg.fraction))
      if (prod === undefined && withheldShape(seg)) { out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'withheld' }); continue }
      // Painted since the upper-segment registration (stage2-scale-upper-segment, 2ee3735a): an upper-segment cell of the
      // num-segment shape, withheld when the searches ran; evaluated as then. A painted fill that is not Excel's is a fault.
      if (prod !== undefined && seg.colours === 3 && seg.segment === 1 && withheldShape(seg)) {
        if (prod !== excel) { out.faults.push(`${c.id}:${ref} production ${prod} is not Excel's ${excel}`); continue }
        out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'withheld' })
        continue
      }
      // Painted since the literal-lower registration (stage2-scale-literal-lower, a3f85309): a literal-num LOWER-segment cell of
      // the num-segment shape, withheld when the searches ran; evaluated as then. A painted fill that is not Excel's is a fault.
      if (prod !== undefined && seg.segment === 0 && withheldShape(seg)) {
        if (prod !== excel) { out.faults.push(`${c.id}:${ref} production ${prod} is not Excel's ${excel}`); continue }
        out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'withheld' })
        continue
      }
      // Withheld only by the later upper-segment whole-product rule (owner ruling A): before it, production painted the
      // truncation here; the cell is evaluated as it was when the search ran.
      // Painted at index-position since the upper-segment registration (2ee3735a): evaluated as when the searches ran; a painted
      // fill that is not Excel's is a fault.
      if (prod !== undefined && prod !== truncate && seg.colours === 3 && seg.segment === 1 && upperSegmentWholeProduct(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction) && prod !== excel) { out.faults.push(`${c.id}:${ref} production ${prod} is not Excel's ${excel}`); continue }
      if ((prod === undefined || prod === excel) && seg.colours === 3 && seg.segment === 1 && upperSegmentWholeProduct(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction)) {
        if (truncate === excel) out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'confirmed' })
        else out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: 'withheld' })
        continue
      }
      // Withheld by the later guard-boundary rule (owner ruling after stage2-scale-upper-window), or painted at raw floor
      // since the raw-floor registration (stage2-scale-raw-floor, ce3f701d): a 2-colour or 3-colour lower-segment cell
      // production painted at its guarded truncation before; evaluated as it was when the searches ran. A painted fill that
      // is not Excel's is a harness fault.
      if (prod !== truncate && (seg.colours === 2 || seg.segment === 0) && guardBoundaryProduct(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction)) {
        if (prod !== undefined && prod !== excel) { out.faults.push(`${c.id}:${ref} production ${prod} is not Excel's ${excel}`); continue }
        out.cells.push({ caseId: c.id, ref, segment: seg, excel, status: truncate === excel ? 'confirmed' : 'withheld' })
        continue
      }
      if (prod === undefined) { exclude('interpolated, then withheld by a later gate (composition)', true); continue }
      if (prod !== truncate) { exclude(`painted ${prod} by another rule over the scale`, false); continue }
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

export { CHANNELS, FRACTIONS, SWEEP_K, SWEEP_N, type Candidate } from './num-segment-formulas'
/** Every candidate (src/xlsx/cf/numSegmentCandidates.ts, the one implementation production's consensus evaluates too). */
export const CANDIDATES: readonly Candidate[] = numSegmentCandidates()
const clamp01 = (f: number): number => Math.max(0, Math.min(1, f))
const channelsOf = (argb: number): number[] => [(argb >>> 24) & 0xff, (argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]
export const BASELINE_ID = 'exact/step-truncate'
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

// ---------------------------------------------------------------------------------------------------------------------
// The confirmation case (stage2-scale-num-segment): its cells as segments, and the survivors' classes on them.
// ---------------------------------------------------------------------------------------------------------------------

/** The segment production evaluates at one cell of a case rule (thresholds ascending, kinds by threshold position). */
function ruleSegment(rule: S2NumSegmentRule, row: number): CfScaleSegment {
  const v = s2NumSegmentValueAt(rule, row), t = rule.thresholds
  const seg = t.length === 2 || v <= t[1]! ? 0 : 1
  const kind = (x: { type: string }): CfScaleSegment['loKind'] => x.type as CfScaleSegment['loKind']
  // The midpoint-above-max rule's cfvo 1 and 2 swap with their thresholds.
  const byPos = rule.cfvo.length === 3 && rule.cfvo[1]!.val !== undefined && rule.cfvo[2]!.val !== undefined && Number(rule.cfvo[1]!.val) > Number(rule.cfvo[2]!.val) ? [rule.cfvo[0]!, rule.cfvo[2]!, rule.cfvo[1]!] : rule.cfvo
  return { value: v, lo: t[seg]!, hi: t[seg + 1]!, fraction: clamp01((v - t[seg]!) / (t[seg + 1]! - t[seg]!)), c0: Number.parseInt(rule.colors[seg]!, 16) >>> 0, c1: Number.parseInt(rule.colors[seg + 1]!, 16) >>> 0, colours: rule.colors.length as 2 | 3, segment: seg as 0 | 1, loKind: kind(byPos[seg]!), hiKind: kind(byPos[seg + 1]!), thresholds: [...t] }
}

/** Each designated cell of the case's evidence rules for an id, as the segment production evaluates (from the rules, never the readings). */
export function caseSegments(capabilityId: string): Array<{ ref: string; segment: CfScaleSegment }> {
  return s2NumSegmentRules().filter(r => r.capabilityId === capabilityId).flatMap(r => (r.designatedRows ?? r.rows).map(row => ({ ref: `${r.col}${row}`, segment: ruleSegment(r, row) })))
}

/** The survivors of the old captures at a scope (pinned by the candidates test). */
export function survivorsAt(cells: readonly SegmentCell[], scope: string): string[] {
  return evaluateCandidates(cells).filter(r => r.scope === scope && r.mismatches.length === 0).map(r => r.id)
}

/**
 * Fix round 1 (R1): a class's reading name, from its representative member (the first, in the enumeration order, of the
 * forms the case's independent model reads): the percent round trip; else the fixed-point grid at its smallest N; else a
 * scaled round trip. A class holding none of these would need a model (an error, never a silent name).
 */
function classNameOf(members: readonly string[]): string {
  if (members.includes('percent-roundtrip/step-truncate-unguarded')) return 'percent-roundtrip'
  const ns = members.filter(id => /^sweep-fixed-\d+$/.test(id)).map(id => Number(id.slice('sweep-fixed-'.length)))
  if (ns.length > 0) return `fixed-${Math.min(...ns)}`
  const rt = members.find(id => /^sweep-roundtrip(-reciprocal|-after)?-\d+$/.test(id))
  if (rt) return rt.slice('sweep-'.length)
  // A class of grid formulations only: named by its first member, `fixed-floor-N/<channel>` (the model reads that form).
  if (/^fixed-floor-\d+\/(step-truncate|step-truncate-unguarded|convex-toward-start|value-toward-start)$/.test(members[0]!)) return members[0]!
  throw new Error(`num-segment: class ${members.slice(0, 3).join(', ')}… has no modelled representative`)
}

/**
 * Fix round 1 (R1): the classes of the survivors (3-colour literal-num scope) on the case's literal-num cells (every
 * evidence rule of scale:num-segment-fraction, separators included), each named by its representative, plus `truncate`
 * (production's arithmetic, a class of one). The committed membership is scripts/cf-refs/num-segment-classes.ts.
 */
export function computeNumSegmentClasses(cells: readonly SegmentCell[]): Record<string, string[]> {
  const survivors = survivorsAt(cells, '3-colour-literal-num')
  const segs = caseSegments(S2_NUM_SEGMENT_CAP)
  const groups = new Map<string, string[]>()
  for (const id of survivors) { const c = CANDIDATES_BY_ID.get(id)!; const k = segs.map(s => hex(c.colour(s.segment))).join(','); groups.set(k, [...(groups.get(k) ?? []), id]) }
  const out: Record<string, string[]> = { truncate: [BASELINE_ID] }
  for (const members of groups.values()) {
    const name = classNameOf(members)
    if (out[name]) throw new Error(`num-segment: two classes named ${name}`)
    out[name] = members
  }
  return out
}
const CANDIDATES_BY_ID = new Map(CANDIDATES.map(c => [c.id, c]))

/** Writes the committed class membership (src) and the reading names (catalogue), both derived from `computeNumSegmentClasses`. */
export function writeNumSegmentClasses(classes: Record<string, string[]>, rootDir: string = REPO_ROOT): void {
  const names = Object.keys(classes)
  writeFileSync(resolve(rootDir, 'scripts/cf-refs/num-segment-classes.ts'), [
    '// Generated by `npx tsx scripts/cf-refs/num-segment-candidates.ts --write-classes` (do not edit).',
    '/**',
    ' * The class of each pre-registered scale:num-segment-fraction reading: the finite list of candidates',
    ' * (scripts/cf-refs/num-segment-formulas.ts ids) that survive every earlier capture and predict the reading\'s fills on every',
    ' * literal-num cell of stage2-scale-num-segment. The capture matched none of them, so nothing is registered and no production',
    ' * code reads this; it documents the search (pinned by tests).',
    ' */',
    `export const CF_NUM_SEGMENT_CLASSES: Readonly<Record<string, readonly string[]>> = Object.freeze({`,
    ...names.map(n => `  ${JSON.stringify(n)}: Object.freeze(${JSON.stringify(classes[n])}),`),
    '})',
    '',
  ].join('\n'))
  writeFileSync(resolve(rootDir, 'scripts/cf-refs/num-segment-readings.ts'), [
    '// Generated by `npx tsx scripts/cf-refs/num-segment-candidates.ts --write-classes` (do not edit).',
    '/** The literal-num readings of stage2-scale-num-segment: `truncate`, then one per class of surviving candidates (scripts/cf-refs/num-segment-classes.ts). */',
    `export const S2_NUM_SEGMENT_READINGS: readonly string[] = ${JSON.stringify(names)}`,
    '',
  ].join('\n'))
}

/** A test domain of 3-colour literal-num cells (eight bound sets, P/Q in both orders and the stock palette, 59 values per segment). */
export function consensusDomain(): CfScaleSegment[] {
  const P = 0xff14f028, Q = 0xffe61efa
  const out: CfScaleSegment[] = []
  for (const t of [[0, 30, 60], [7, 217, 637], [0, 1000, 2000], [0, 8, 16], [10, 17, 24], [-40, 15, 100], [0, 100, 200], [1, 6, 11]])
    for (const pal of [[P, Q, P], [Q, P, Q], [0xfff8696b, 0xffffeb84, 0xff63be7b]]) for (const seg of [0, 1] as const) {
      const lo = t[seg]!, hi = t[seg + 1]!
      for (let i = 1; i < 60; i++) { const v = Number((lo + ((hi - lo) * i) / 60).toFixed(4)); out.push({ value: v, lo, hi, fraction: (v - lo) / (hi - lo), c0: pal[seg]!, c1: pal[seg + 1]!, colours: 3, segment: seg, loKind: 'num', hiKind: 'num', thresholds: t }) }
    }
  return out
}
/** The share of a domain where a class's members disagree (consensus withholds there). */
export function consensusWithheldShare(members: readonly string[], domain: readonly CfScaleSegment[]): number {
  const cs = members.map(id => CANDIDATES_BY_ID.get(id)!)
  let n = 0
  for (const s of domain) { const a = cs[0]!.colour(s); if (cs.some(c => c.colour(s) !== a)) n++ }
  return n / domain.length
}

/**
 * Fix round 1 (R1): the greedy search that chose `S2_NUM_SEGMENT.separators` (slow: run with `--choose-separators`).
 * Each separator is a lower value v of the 7 / 217 / 637 rules with its upper twin 217 + 2·(v − 7), placed in both colour
 * orders (four cells in one configuration). Starting from the forced dyadic eighth (33.25), each round adds the pool value
 * that most lowers the class-size-weighted share of `consensusDomain()` (sampled at 59 values per segment) that consensus
 * would withhold, subject to every pair of classes differing at 0 or >= 3 cells of each configuration; it stops after
 * `rounds` or when nothing lowers the share.
 */
export function chooseSeparators(cells: readonly SegmentCell[], rounds = 12, forced: readonly number[] = [33.25]): number[] {
  const survivors = survivorsAt(cells, '3-colour-literal-num')
  const cs = survivors.map(id => CANDIDATES_BY_ID.get(id)!)
  const P = 0xff14f028, Q = 0xffe61efa, T = [7, 217, 637]
  const mk = (val: number, lo: number, hi: number, c0: number, c1: number, segment: 0 | 1): CfScaleSegment => ({ value: val, lo, hi, fraction: (val - lo) / (hi - lo), c0, c1, colours: 3, segment, loKind: 'num', hiKind: 'num', thresholds: T })
  const place = (v: number): CfScaleSegment[][] => { const u = 217 + 2 * (v - 7); return [[mk(v, 7, 217, P, Q, 0), mk(u, 217, 637, Q, P, 1)], [mk(v, 7, 217, Q, P, 0), mk(u, 217, 637, P, Q, 1)]] }
  const base = S2_NUM_SEGMENT.fractions.map(([p, q]) => 7 + (210 * p) / q)
  const pool: number[] = []
  for (const k of [0, 1, 2, 3, 5, 8, 13, 21, 26, 34, 52, 70, 89, 105, 140, 175, 200]) for (const d of [0.001, 0.002, 0.004, 0.008, 0.016, 0.03, 0.06, 0.12, 0.25, 0.4, 0.6, 0.8]) { const v = Number((7 + k + d).toFixed(3)); if (v < 217) pool.push(v) }
  for (let v = 8; v <= 216; v++) pool.push(v)
  pool.push(33.25, 85.75, 8.26, 9.1, 20.125, 32.375)
  const cl = new Map<number, number[][][]>()
  for (const v of [...base, ...pool]) cl.set(v, place(v).map(c => c.map(x => cs.map(k => k.colour(x)))))
  const lab = (v: number, i: number): string => cl.get(v)!.flat().map(pl => pl[i]).join(':')
  const domain = consensusDomain()
  const dom = cs.map(k => domain.map(s => k.colour(s)))
  const partition = (vs: readonly number[]): number[][] => { const g = new Map<string, number[]>(); survivors.forEach((_, i) => { const key = vs.map(v => lab(v, i)).join('|'); g.set(key, [...(g.get(key) ?? []), i]) }); return [...g.values()] }
  const withhold = (x: readonly number[]): number => { let n = 0; for (let k = 0; k < domain.length; k++) { const c = dom[x[0]!]![k]; for (const i of x) if (dom[i]![k] !== c) { n++; break } } return n / domain.length }
  const score = (vs: readonly number[]): number => partition(vs).reduce((a, x) => a + withhold(x) * x.length, 0) / survivors.length
  const ok = (sep: readonly number[]): boolean => {
    const reps = partition([...base, ...sep]).map(x => x[0]!)
    for (let a = 0; a < reps.length; a++) for (let b = a + 1; b < reps.length; b++) {
      const A = reps[a]!, B = reps[b]!, n = [0, 0, 0]
      for (const v of base) for (let c = 0; c < 2; c++) for (const pl of cl.get(v)![c]!) if (pl[A] !== pl[B]) n[c]!++
      for (const v of sep) for (const cc of cl.get(v)!) for (const pl of cc) if (pl[A] !== pl[B]) n[2]!++
      if (n.some(x => x > 0 && x < 3)) return false
    }
    return true
  }
  const sep = [...forced]
  let cur = score([...base, ...sep])
  for (let round = 0; round < rounds; round++) {
    let best: { v: number; s: number } | undefined
    for (const v of pool) {
      if (sep.includes(v) || base.includes(v)) continue
      const s = score([...base, ...sep, v])
      if (s >= cur - 1e-9 || (best && s >= best.s) || !ok([...sep, v])) continue
      best = { v, s }
    }
    if (!best) break
    sep.push(best.v); cur = best.s
  }
  return sep
}

async function main(): Promise<void> {
  const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const { cells, excluded, faults } = await collectSegmentCells(manifest)
  if (process.argv.includes('--choose-separators')) { console.log(`separators: ${JSON.stringify(chooseSeparators(cells))}`); return }
  const byCase = new Map<string, { confirmed: number; withheld: number }>()
  for (const c of cells) { const e = byCase.get(c.caseId) ?? { confirmed: 0, withheld: 0 }; e[c.status]++; byCase.set(c.caseId, e) }
  console.log(`## Cells\n\n${cells.length} captured colour-scale cells (${cells.filter(c => c.status === 'confirmed').length} confirmed, ${cells.filter(c => c.status === 'withheld').length} withheld as the num-segment shape); ${excluded.length} excluded; ${faults.length} harness faults.\n`)
  console.log('| Case | Confirmed | Withheld (num-segment) |\n|---|---|---|')
  for (const [id, e] of byCase) console.log(`| \`${id}\` | ${e.confirmed} | ${e.withheld} |`)
  for (const group of ['a', 'b'] as const) {
    const exReasons = new Map<string, number>()
    for (const e of excluded.filter(x => x.group === group)) { const k = `${e.caseId}: ${e.reason}${group === 'b' ? (e.threeColour ? ' [under a 3-colour scale]' : ' [2-colour scales only]') : ''}`; exReasons.set(k, (exReasons.get(k) ?? 0) + 1) }
    console.log(`\nExcluded, group ${group} (${group === 'a' ? 'inapplicable' : 'numeric, Excel-filled, dropped only by a production gate'}): ${excluded.filter(x => x.group === group).length}`)
    for (const [r, n] of exReasons) console.log(`- ${r} (${n})`)
  }
  for (const f of faults) console.log(`FAULT ${f}`)
  const interior = (c: SegmentCell): boolean => c.segment.value > c.segment.lo && c.segment.value < c.segment.hi
  const classOf = (c: SegmentCell): string => `${c.segment.colours}-colour ${c.segment.loKind}/${c.segment.hiKind}`
  const classes = new Map<string, { cells: number; interior: number; whole: number; lowered: number }>()
  for (const c of cells) {
    const e = classes.get(classOf(c)) ?? { cells: 0, interior: 0, whole: 0, lowered: 0 }
    e.cells++
    if (interior(c)) {
      e.interior++
      const w = fractionWindow(c), f = FRACTIONS.exact!(c.segment)
      if (w.lo === f) e.whole++
      if (f >= w.hi) e.lowered++
    }
    classes.set(classOf(c), e)
  }
  console.log('\n| Segment class | Cells | Strictly inside | Whole product, drawn at truncation (rejects any lowered fraction) | Drawn below truncation |\n|---|---|---|---|---|')
  for (const [k, e] of [...classes].sort()) console.log(`| ${k} | ${e.cells} | ${e.interior} | ${e.whole} | ${e.lowered} |`)
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

  console.log('\n## Classes on the case\n')
  const numClasses = computeNumSegmentClasses(cells)
  if (process.argv.includes('--write-classes')) writeNumSegmentClasses(numClasses)
  const domain = consensusDomain()
  const shares = Object.fromEntries(Object.entries(numClasses).map(([n, m]) => [n, consensusWithheldShare(m, domain)]))
  const names = Object.keys(numClasses)
  console.log(`${names.length} readings (truncate + ${names.length - 1} classes of the ${survivorsAt(cells, '3-colour-literal-num').length} survivors) on ${caseSegments(S2_NUM_SEGMENT_CAP).length} literal-num cells; consensus domain ${domain.length} cells.\n`)
  const sizes = names.slice(1).map(n => numClasses[n]!.length)
  const weighted = names.slice(1).reduce((a, n) => a + shares[n]! * numClasses[n]!.length, 0) / sizes.reduce((a, b) => a + b, 0)
  const sorted = names.slice(1).map(n => shares[n]!).sort((a, b) => a - b)
  console.log(`Class sizes: ${[...sizes].sort((a, b) => b - a).join(', ')}.`)
  console.log(`Domain withheld under consensus: size-weighted mean ${(weighted * 100).toFixed(2)}%, median ${(sorted[Math.floor(sorted.length / 2)]! * 100).toFixed(2)}%, max ${(sorted.at(-1)! * 100).toFixed(2)}%; ${sorted.filter(x => x === 0).length} classes withhold nothing.\n`)
  console.log('| Reading | Members | Domain withheld | Members (first six) |\n|---|---|---|---|')
  for (const n of names) console.log(`| \`${n}\` | ${numClasses[n]!.length} | ${(shares[n]! * 100).toFixed(2)}% | ${numClasses[n]!.slice(0, 6).join(', ')}${numClasses[n]!.length > 6 ? ', …' : ''} |`)
  const stat = survivorsAt(cells, '3-colour')
  console.log(`\nscale:num-segment-fraction-statistical: the ${stat.length} survivors over every 3-colour segment (${stat.join(', ')}) are the percent round trip; readings truncate, percent-roundtrip.\n`)

  console.log('## Predictions at the decisive cells\n')
  const lines: string[] = []
  for (const cfg of s2NumSegmentLayout().configurations) {
    const hs = Object.keys(cfg.hypotheses)
    const decisive = cfg.refs.filter(ref => new Set(hs.map(h => cfg.hypotheses[h]![ref])).size > 1)
    let min = Infinity
    for (let i = 0; i < hs.length; i++) for (let j = i + 1; j < hs.length; j++) { const n = cfg.refs.filter(r => cfg.hypotheses[hs[i]!]![r] !== cfg.hypotheses[hs[j]!]![r]).length; if (n > 0) min = Math.min(min, n) }
    console.log(`\`${cfg.id}\` (${cfg.capabilityId}${cfg.role ? `, ${cfg.role}` : ''}): ${cfg.refs.length} cells, ${decisive.length} decisive; smallest separation of two differing readings ${min === Infinity ? '—' : min}.`)
    lines.push(`## \`${cfg.id}\` (${cfg.capabilityId}${cfg.role ? `, ${cfg.role}` : ''})`, '', '| Cell | Predicted token: readings |', '|---|---|')
    for (const ref of cfg.refs) {
      const by = new Map<string, string[]>()
      for (const h of hs) by.set(cfg.hypotheses[h]![ref]!, [...(by.get(cfg.hypotheses[h]![ref]!) ?? []), h])
      lines.push(`| ${ref} | ${[...by].map(([t, rs]) => `${t}: ${rs.length === hs.length ? 'every reading' : rs.join(', ')}`).join('; ')} |`)
    }
    lines.push('')
  }
  const { twins } = s2NumSegmentLayout()
  lines.push('## Twin line (2-colour num 0 / 210, P to Q)', '', '| Twin | k | Fill (P + k·(1, −1, 1)) |', '|---|---|---|', ...twins.map(t => `| ${t.ref} | ${t.k} | ${hex((0xff14f028 + t.k * 0x10000 - t.k * 0x100 + t.k) >>> 0)} |`), '')
  if (process.argv.includes('--write-predictions')) writeFileSync(resolve(REPO_ROOT, 'docs/validation/xlsx-cf-num-segment-predictions.md'), ['# stage2-scale-num-segment: every reading\'s prediction at every designated cell', '', 'Generated by `npx tsx scripts/cf-refs/num-segment-candidates.ts --write-predictions` from the catalogue (`scaleReadings`); a `twin:` token is the twin line\'s cell listed at the end.', '', ...lines].join('\n'))
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch(err => { console.error(err); process.exit(1) })
}
