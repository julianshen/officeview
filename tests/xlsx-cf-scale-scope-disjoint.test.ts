// @vitest-environment node
/**
 * The scope-disjoint follow-up (brief 2026-11-21, Ruling F-2): stage2-scale-scope-disjoint, discriminating evidence for
 * scale:scope-disjoint (a colour scale over several ranges).
 *
 * stage2-scale-scope cannot discriminate it: its B2:B4 D2:D4 rule overlaps two others (every non-hidden cell withheld,
 * S2-I) and both rectangles hold the same values, so union and per-range statistics paint alike. This case version holds
 * two non-overlapping disjoint scales (2-colour min/max, 3-colour min / percentile 50 / max) whose rectangles hold other
 * values, a single-range negative control and literal num twins, one per predicted rendering. Its readings (union,
 * per-range, leftmost-range, rightmost-range) are pre-registered in the catalogue and judged by the Stage 5 configuration
 * evaluator over Excel's DisplayFormat fills.
 *
 * It has no capture yet. What is checked here, before any capture:
 * 1. production already computes union statistics for a multi-range scale (only the gate withholds it);
 * 2. the case design: ownership, isolation, size, and >= 2 decisive cells per rectangle for every pair of readings;
 * 3. each reading, planted on a synthetic observation by an independent model, comes back as the one match;
 * 4. the audit gate end to end (real verifyCfCase and auditStage2Manifest on a synthetic measured capture): only a
 *    discriminating `union` (production's reading) verifies, and only once the registry entry is gone;
 * 5. the rehearsal of the publication: with scale:scope-disjoint held, production agrees with every committed Excel
 *    capture (Stages 1-4 unchanged, Stage 5 records with no disagreement or unjustified refusal) and paints the planted
 *    union observation of this case exactly.
 * This checks the family and production, not Excel.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import {
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE2_NONDISCRIMINATING,
  CF_STAGE2_OWNERS,
  S2_SCOPE_DISJOINT,
  S2_SCOPE_DISJOINT_CAP,
  S2_SCOPE_DISJOINT_CONTROL,
  S2_SCOPE_DISJOINT_ID,
  S2_SCOPE_DISJOINT_READINGS,
  S2_SCOPE_DISJOINT_TWIN,
  s2ScopeDisjointBounds,
  s2ScopeDisjointStatValues,
  type CfManifest,
  type CfOracleCase,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  capabilityOwnerCase,
  computeWorksheetRangeGeometry,
  evaluateFormulaCapabilities,
  formulaControlsOf,
  isCellInSqref,
  readZipPartsSync,
  stage2ReadingsObservationOf,
  verifyCfCase,
} from '../scripts/cf-refs/verify'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { interpolateScaleColor } from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { XlsxSheet } from '../src/xlsx/types'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_SCOPE_DISJOINT_ID
const CAP = S2_SCOPE_DISJOINT_CAP
const READINGS = S2_SCOPE_DISJOINT_READINGS
const HEAVY = 300_000

const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === ID)
  if (!c) throw new Error(`${ID} is not in the manifest (generate it with generate.ts --case ${ID})`)
  return c
}
const partsOf = (c: CfOracleCase, file = c.saved): Map<string, Uint8Array> => readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, file))))!
const valuesOf = (c: CfOracleCase): Map<string, number> => {
  const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
  const out = new Map<string, number>()
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"><v>([^<]+)<\/v><\/c>/g)) out.set(m[1]!, Number(m[2]))
  return out
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
const rc = (ref: string): { row: number; col: number } => ({ row: rowOf(ref) - 1, col: [...colOf(ref)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 })
const FX = S2_SCOPE_DISJOINT
const DATA_REFS = (cols: readonly string[]): string[] => cols.flatMap(col => FX.rows.map(r => `${col}${r}`))

// ---- an independent model of Excel's scale colour (Rulings S2-B: sRGB, truncate; S2-E: clamped beyond the bounds) ----
const COLOURS: Record<2 | 3, readonly number[]> = { 2: [0xfff8696b, 0xff63be7b], 3: [0xfff8696b, 0xffffeb84, 0xff63be7b] }
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
function modelColour(v: number, bounds: readonly number[], colours: 2 | 3): string {
  const cs = COLOURS[colours]
  if (v <= bounds[0]!) return hex(cs[0]!)
  if (v >= bounds[bounds.length - 1]!) return hex(cs[cs.length - 1]!)
  const seg = colours === 2 || v < bounds[1]! ? 0 : 1
  const lo = bounds[seg]!, hi = bounds[seg + 1]!
  return hex(interpolateScaleColor(cs[seg]!, cs[seg + 1]!, (v - lo) / (hi - lo)))
}

/** The twin marker (single range, num bounds) that formats `ref`, if any. */
const twinOf = (c: CfOracleCase, ref: string) => c.markers.find(m => m.capabilityId === S2_SCOPE_DISJOINT_TWIN && isCellInSqref(ref, m.sqref))

/** Planted DisplayFormat fills: designated cells under `reading`, twins under their own literal rule, controls none. */
function plantedFills(c: CfOracleCase, reading: string): Map<string, string> {
  const values = valuesOf(c)
  const out = new Map<string, string>()
  for (const e of c.expected.cells) {
    const v = values.get(e.ref)!
    if (!e.matched) { out.set(e.ref, 'FFFFFFFF'); continue }
    const sc = FX.scales.find(s => (s.cols as readonly string[]).includes(colOf(e.ref)))
    if (sc) { out.set(e.ref, modelColour(v, s2ScopeDisjointBounds(s2ScopeDisjointStatValues(sc, reading, sc.cols.indexOf(colOf(e.ref)) as 0 | 1), sc.colours), sc.colours)); continue }
    if (colOf(e.ref) === FX.control.col) { out.set(e.ref, modelColour(v, s2ScopeDisjointBounds(FX.control.values, 2), 2)); continue }
    const t = twinOf(c, e.ref)!
    out.set(e.ref, modelColour(v, t.scale!.cfvo.map(x => Number(x.val)), t.scale!.colors.length as 2 | 3))
  }
  return out
}
const plantedSidecar = (c: CfOracleCase, reading: string): Record<string, unknown> => ({
  observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: plantedFills(c, reading).get(e.ref) })),
})
const judge = (c: CfOracleCase, sidecar: Record<string, unknown>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(sidecar))

async function sheetOf(c: CfOracleCase): Promise<XlsxSheet> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
  return doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
}
/** Production's fill at each ref (resident paint under whatever evidence is in force). */
async function productionFills(c: CfOracleCase, refs: readonly string[]): Promise<Map<string, string | undefined>> {
  const paint = await prepareCfPaint(await sheetOf(c), {})
  const out = new Map<string, string | undefined>()
  for (const ref of refs) { const { row, col } = rc(ref); out.set(ref, paint?.layer.styleAt(row, col, undefined)?.fillColor) }
  paint?.dispose()
  return out
}
const WITHOUT_DISJOINT = (): string[] => [...VERIFIED_CAPABILITIES].filter(id => id !== CAP)
/** The capture's DisplayFormat fills (e448d831). */
const excelFills = (c: CfOracleCase): Map<string, string> => {
  const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
  return new Map(sc.observedCells.map(o => [o.ref, o.fillRgb]))
}
/**
 * The withheld shape of the finding (a87121e8, `numSegmentUnmeasured`), recomputed from the twin's marker: a 3-colour scale
 * whose value lies strictly inside a segment bounded by two literal num cfvo, where f*2^20 is not an integer.
 */
function inNumSegmentShape(c: CfOracleCase, ref: string): boolean {
  const t = twinOf(c, ref)
  if (!t || t.scale!.colors.length !== 3) return false
  const [a, b, d] = t.scale!.cfvo.map(x => Number(x.val))
  const v = valuesOf(c).get(ref)!
  const [lo, hi] = v <= b! ? [a!, b!] : [b!, d!]
  return v > lo && v < hi && !Number.isInteger(((v - lo) / (hi - lo)) * 1048576)
}

describe('the capture (e448d831): union, and production against Excel', () => {
  it('measured: verifyCfCase matches union; the twins and the control are never a question; nothing refuses', () => {
    const c = caseOf()
    expect(c.status).toBe('measured')
    const res = verifyCfCase(c, { rootDir: REPO_ROOT })
    expect(res.kind, res.kind === 'refusal' ? res.reason : '').toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [CAP]: 'union' })
    expect(res.verifiedCapabilities).toEqual([CAP])
    expect(res.nonDiscriminatingCapabilities).toEqual([S2_SCOPE_DISJOINT_CONTROL, S2_SCOPE_DISJOINT_TWIN])
    expect(res.refusedCapabilities).toEqual([])
  })

  it('every designated cell (both disjoint scales and the control) paints exactly Excel\'s fill, its union twin\'s', async () => {
    const c = caseOf()
    const excel = excelFills(c)
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    for (const cfg of c.scaleReadings!) for (const ref of cfg.refs) {
      expect(fills.get(ref), `${cfg.id} ${ref}`).toBe(excel.get(ref))
      expect(excel.get(ref), `${cfg.id} ${ref} vs its union twin`).toBe(excel.get(cfg.hypotheses.union![ref]!.slice('twin:'.length)))
    }
    for (const e of c.expected.cells.filter(x => !x.matched)) expect(fills.get(e.ref), e.ref).toBeUndefined()
  }, HEAVY)

  it('every twin paints exactly Excel\'s fill, except the num-segment shape, which is withheld (T4 and V3 are where Excel differs)', async () => {
    const c = caseOf()
    const excel = excelFills(c)
    const twins = c.expected.cells.filter(e => twinOf(c, e.ref)).map(e => e.ref)
    const fills = await productionFills(c, twins)
    const withheld = twins.filter(ref => inNumSegmentShape(c, ref))
    expect(withheld).toEqual(['R3', 'R4', 'R5', 'S2', 'S3', 'S4', 'T3', 'T4', 'U3', 'V3', 'W4'])
    // U4 (num 30/60/100, value 70, f = 1/4, upper segment, step 39 whole) is withheld since stage2-scale-num-segment
    // (the upper-segment whole-product withhold, owner ruling A); Excel drew it at truncation.
    for (const ref of twins) {
      if (withheld.includes(ref) || ref === 'U4') expect(fills.get(ref), ref).toBeUndefined()
      else expect(fills.get(ref), ref).toBe(excel.get(ref))
    }
    // The finding: with the shape painted (production's interpolation), T4 and V3 are one unit lower per RGB channel than Excel.
    const planted = plantedFills(c, 'union')
    expect(withheld.filter(ref => planted.get(ref) !== excel.get(ref))).toEqual(['T4', 'V3'])
    expect([excel.get('T4'), planted.get('T4'), excel.get('V3'), planted.get('V3')]).toEqual(['FFCCDD82', 'FFCBDC81', 'FF98CE7F', 'FF97CD7E'])
  }, HEAVY)

  it('without scale:scope-disjoint the multi-range cells are withheld, announced unavailable; with it they are not', async () => {
    const c = caseOf()
    expect(VERIFIED_CAPABILITIES.has(CAP)).toBe(true)
    const multi = c.scaleReadings!.filter(x => x.role !== 'negative-control').flatMap(x => x.refs)
    const fills = await withExactCfEvidence(WITHOUT_DISJOINT(), () => productionFills(c, multi))
    for (const ref of multi) expect(fills.get(ref), ref).toBeUndefined()
    // The text scope (the announcement path every view shares; pagedSource builds the same service) says so too.
    const sheet = await sheetOf(c)
    const describeAt = async (ref: string): Promise<string | undefined> => {
      const { row, col } = rc(ref)
      const budget = { reserve: () => {}, release: () => {}, retain: (_t: unknown, _b: number, commit: () => void) => commit(), settle: () => {} }
      const scope = await prepareCfTextScope(sheet, { rects: [{ r0: row, r1: row, c0: col, c1: col }], signal: new AbortController().signal, locale: 'en-US', budget: budget as never })
      try { return scope?.describe(row, col)?.text } finally { scope?.dispose() }
    }
    expect(await withExactCfEvidence(WITHOUT_DISJOINT(), () => describeAt('B3'))).toBe('Conditional formatting unavailable')
    expect(await withExactCfEvidence(WITHOUT_DISJOINT(), () => describeAt('H3'))).toBe('Conditional formatting unavailable')
    expect(await describeAt('B3')).not.toBe('Conditional formatting unavailable')
    expect(await describeAt('H3')).not.toBe('Conditional formatting unavailable')
    // A withheld twin of the num-segment shape is announced unavailable too.
    expect(await describeAt('T4')).toBe('Conditional formatting unavailable')
  }, HEAVY)
})

describe('stage2-scale-scope-disjoint: the case', () => {
  it('is appended after the earlier cases and owns scale:scope-disjoint; the old case keeps its marker and capture as a cross-check', () => {
    // Appended last at its time; the num-segment follow-up (stage2-scale-num-segment) is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-5)).toEqual([ID, 'stage2-scale-num-segment', 'stage2-scale-upper-whole', 'stage2-scale-upper-window', 'stage2-scale-raw-floor'])
    expect(CF_STAGE2_OWNERS[CAP]).toBe(ID)
    expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === CAP).map(e => e.caseId)).toEqual([ID])
    expect(capabilityOwnerCase(CAP, manifest)!.id).toBe(ID)
    const old = manifest.cases.find(x => x.id === 'stage2-scale-scope')!
    expect(old.status).toBe('measured')
    expect(old.markers.some(m => m.capabilityId === CAP)).toBe(true)
    // One registry entry per distinct id; the comparison ids are registered, never a question.
    const ids = CF_STAGE2_CAPABILITIES.map(e => e.capabilityId)
    expect(new Set(ids).size).toBe(ids.length)
    for (const id of [S2_SCOPE_DISJOINT_TWIN, S2_SCOPE_DISJOINT_CONTROL]) expect(CF_STAGE2_NONDISCRIMINATING[id], id).toBeDefined()
    // Measured (e448d831): the registry entry is gone, removed by the publication.
    expect(CF_STAGE2_NONDISCRIMINATING[CAP]).toBeUndefined()
    expect(caseOf().status).toBe('measured')
  })

  it('the rules: two disjoint scales, a single-range control and num twins; nothing overlaps, nothing is hidden or merged', () => {
    const c = caseOf()
    expect(c.markers.filter(m => m.capabilityId !== S2_SCOPE_DISJOINT_TWIN).map(m => [m.capabilityId, m.sqref, m.scale!.cfvo.map(x => x.type + (x.val ?? '')).join(',')])).toEqual([
      [CAP, 'B2:B5 D2:D5', 'min,max'],
      [CAP, 'F2:F5 H2:H5', 'min,percentile50,max'],
      [S2_SCOPE_DISJOINT_CONTROL, 'J2:J5', 'min,max'],
    ])
    for (const m of c.markers.filter(x => x.capabilityId === S2_SCOPE_DISJOINT_TWIN)) expect(m.scale!.cfvo.every(x => x.type === 'num'), m.sqref).toBe(true)
    const cells = c.expected.cells.map(e => e.ref)
    for (const ref of cells) expect(c.markers.filter(m => isCellInSqref(ref, m.sqref)).length, ref).toBeLessThanOrEqual(1)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    // Every designated and twin cell holds a number; no rule's range is adjacent to another rectangle of the same rule.
    const values = valuesOf(c)
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
  })

  it('small: 88 cells (20 designated, 48 twins, 20 controls), under 100', () => {
    const c = caseOf()
    expect(c.expected.cells.length).toBe(88)
    expect(c.expected.cells.filter(e => twinOf(c, e.ref)).length).toBe(48)
    expect(c.expected.negativeRefs.length).toBe(20)
    expect(c.range).toBe('A1:W10')
  })

  it('readings: every configuration names the four readings; the 2- and 3-colour configurations replicate per rectangle; J is the negative control', () => {
    const cfgs = caseOf().scaleReadings!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.observable, x.role ?? '', x.replicateBy ?? ''])).toEqual([
      ['disjoint-2colour', CAP, 'fill', '', 'column'], ['disjoint-3colour', CAP, 'fill', '', 'column'], ['single-range-control', CAP, 'fill', 'negative-control', ''],
    ])
    for (const x of cfgs) expect(Object.keys(x.hypotheses), x.id).toEqual([...READINGS])
    // The catalogue's readings are the manifest's (Lesson 14: fixed before capture).
    expect(caseOf().scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
  })

  it('every pair of readings puts >= 3 cells of each rectangle on distinct colours, robust to 0.1 of a unit per channel (pre-capture ruling)', () => {
    const c = caseOf()
    const values = valuesOf(c)
    const channels = (v: number, bounds: readonly number[], colours: 2 | 3): Array<[number, number]> => {
      const cs = COLOURS[colours]
      const ch = (n: number): number[] => [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff]
      if (v <= bounds[0]!) return ch(cs[0]!).map(x => [x, x])
      if (v >= bounds[bounds.length - 1]!) return ch(cs[cs.length - 1]!).map(x => [x, x])
      const seg = colours === 2 || v < bounds[1]! ? 0 : 1
      const f = (v - bounds[seg]!) / (bounds[seg + 1]! - bounds[seg]!)
      const a = ch(cs[seg]!), b = ch(cs[seg + 1]!)
      return a.map((c0, i) => { const d = Math.abs(b[i]! - c0), s = b[i]! >= c0 ? 1 : -1, lo = c0 + s * Math.floor(f * d - 0.1), hi = c0 + s * Math.floor(f * d + 0.1); return [Math.min(lo, hi), Math.max(lo, hi)] as [number, number] })
    }
    const apart = (x: Array<[number, number]>, y: Array<[number, number]>): boolean => x.some((p, i) => p[1] < y[i]![0] || y[i]![1] < p[0])
    for (const sc of FX.scales) {
      for (const [ia, a] of READINGS.entries()) for (const b of READINGS.slice(ia + 1)) {
        for (const i of [0, 1] as const) {
          const refs = DATA_REFS([sc.cols[i]])
          const differing = refs.filter(ref => {
            const v = values.get(ref)!
            return apart(channels(v, s2ScopeDisjointBounds(s2ScopeDisjointStatValues(sc, a, i), sc.colours), sc.colours), channels(v, s2ScopeDisjointBounds(s2ScopeDisjointStatValues(sc, b, i), sc.colours), sc.colours))
          })
          const same = (h: string): string => s2ScopeDisjointStatValues(sc, h, i).join(',')
          // A pair that reads the same statistics on a rectangle cannot differ there; every other pair differs at >= 3 cells.
          if (same(a) === same(b)) expect(differing, `${sc.id} ${sc.cols[i]} ${a}/${b}`).toEqual([])
          else expect(differing.length, `${sc.id} ${sc.cols[i]} ${a}/${b}`).toBeGreaterThanOrEqual(3)
        }
        // Every pair of readings differs somewhere in each configuration.
        expect(sc.cols.some((_, i) => s2ScopeDisjointStatValues(sc, a, i as 0 | 1).join() !== s2ScopeDisjointStatValues(sc, b, i as 0 | 1).join()), `${sc.id} ${a}/${b}`).toBe(true)
      }
    }
  })
})

describe('stage2-scale-scope-disjoint: each reading discriminates on a synthetic observation', () => {
  for (const h of READINGS) {
    it(`planted ${h}: ${h} comes back as the one match`, () => {
      const c = caseOf()
      const ev = judge(c, plantedSidecar(c, h))
      expect([ev[CAP]!.status, ev[CAP]!.matched], ev[CAP]!.reason).toEqual(['discriminating', h])
    })
  }

  it('fails closed: a formatted control, a missing twin observation or a twin that renders nothing refuses', () => {
    const c = caseOf()
    const base = plantedSidecar(c, 'union') as { observedCells: Array<Record<string, unknown>> }
    const edit = (fn: (o: Record<string, unknown>) => Record<string, unknown> | undefined) => ({ observedCells: base.observedCells.map(o => fn({ ...o })).filter(Boolean) })
    expect(judge(c, edit(o => (o.ref === 'B7' ? { ...o, fillRgb: 'FFF8696B' } : o)))[CAP]!.status).toBe('refused')
    expect(judge(c, edit(o => (o.ref === 'L3' ? undefined : o)))[CAP]!.status).toBe('refused')
    expect(judge(c, edit(o => (o.ref === 'L3' ? { ...o, fillRgb: 'FFFFFFFF' } : o)))[CAP]!.status).toBe('refused')
    // A world in which no reading holds is a finding, never fitted.
    expect(judge(c, edit(o => (o.ref === 'B3' ? { ...o, fillRgb: 'FF000000' } : o)))[CAP]!.status).toBe('refused')
  })
})

// ---- a synthetic measured capture through the real verifier and audit (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

/** A temp root holding every Stage 2 case's files, with this case measured on a synthetic capture of `reading`. */
function syntheticRoot(reading: string): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-disjoint-'))
  const copy = (rel: string): void => { if (!existsSync(resolve(REPO_ROOT, rel))) return; mkdirSync(dirname(join(root, rel)), { recursive: true }); copyFileSync(resolve(REPO_ROOT, rel), join(root, rel)) }
  for (const x of manifest.cases.filter(k => k.stage === 2)) {
    copy(x.input); copy(x.saved)
    for (const cap of x.captures) { copy(cap.pngPath); copy(cap.sidecarPath) }
  }
  const c = caseOf()
  const saved = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))
  const input = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input)))
  const geom = computeWorksheetRangeGeometry(readZipPartsSync(saved)!, c.range)!
  const width = Math.round(geom.expectedWidthPt) + 1, height = Math.round(geom.expectedHeightPt) + 1
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, width, height)
  const fills = plantedFills(c, reading)
  for (const e of c.expected.cells) {
    const { row, col } = rc(e.ref)
    const ci = col - geom.minCol, ri = row - geom.minRow
    const x0 = Math.round(geom.colWidthsPt.slice(0, ci).reduce((a, b) => a + b, 0)), x1 = Math.round(geom.colWidthsPt.slice(0, ci + 1).reduce((a, b) => a + b, 0))
    const y0 = Math.round(geom.rowHeightsPt.slice(0, ri).reduce((a, b) => a + b, 0)), y1 = Math.round(geom.rowHeightsPt.slice(0, ri + 1).reduce((a, b) => a + b, 0))
    ctx.fillStyle = `#${fills.get(e.ref)!.slice(2)}`
    ctx.fillRect(x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1)
  }
  const png = new Uint8Array(canvas.toBuffer('image/png'))
  const cap = c.captures[0]!
  mkdirSync(dirname(join(root, cap.pngPath)), { recursive: true })
  writeFileSync(join(root, cap.pngPath), png)
  const at = '2026-11-21T00:00:00.000Z'
  writeFileSync(join(root, cap.sidecarPath), JSON.stringify({
    excelVersion: '16.114', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width, height,
    colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: sha(saved),
    sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256: sha(saved), pngSha256: sha(png), capturedAt: at,
    observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
  }, null, 2))
  const measured: CfOracleCase = {
    ...c,
    status: 'measured',
    reason: undefined,
    resavedByExcel: true,
    resaveReceipt: { openedRepairFree: true, excelVersion: '16.114', resavedAt: at, inputSha256: sha(input), savedSha256: sha(saved) } as CfOracleCase['resaveReceipt'],
    captures: [{ ...cap, pngSha256: sha(png) }],
    expected: { ...c.expected, provenance: 'excel-observed', cells: c.expected.cells.map(e => (e.matched ? { ...e, fillSource: 'excel-display-format' as const } : e)) },
  }
  return { root, manifest: { ...manifest, cases: manifest.cases.map(x => (x.id === ID ? measured : x)) } }
}

describe('the audit gate on a synthetic measured capture (real verifyCfCase and auditStage2Manifest)', () => {
  const withoutEntry = CF_STAGE2_NONDISCRIMINATING
  const withEntry = { ...CF_STAGE2_NONDISCRIMINATING, [CAP]: 'the pre-publication registry entry (rehearsal seam)' }
  const disposition = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }, id: string) => r.perCapability.find(p => p.capabilityId === id)!.disposition
  const beforeCap = [...VERIFIED_STAGE2_CAPABILITIES].filter(id => id !== CAP).sort()

  it('the real capture: scale:scope-disjoint is verified, owned by the new case, and the audit verifies exactly the published Stage 2 set', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    expect(r.perCapability.find(p => p.capabilityId === CAP)).toEqual({ capabilityId: CAP, caseId: ID, disposition: 'verified' })
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...VERIFIED_STAGE2_CAPABILITIES].sort())
    // With the pre-publication registry entry put back (the seam), the same capture is withheld.
    expect(disposition(auditStage2Manifest(manifest, { rootDir: REPO_ROOT, nonDiscriminating: withEntry }), CAP)).toBe('unverified-nondiscriminating')
  }, HEAVY)

  it('planted union: the readings discriminate union; the audit verifies it only once the registry entry is gone (the publication, rehearsed)', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const { root, manifest: m } = syntheticRoot('union')
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [CAP]: 'union' })
    expect(res.verifiedCapabilities).toEqual([CAP])
    expect(res.nonDiscriminatingCapabilities).toEqual([S2_SCOPE_DISJOINT_CONTROL, S2_SCOPE_DISJOINT_TWIN])
    // With the entry still registered: measured, but never verified (the registry entry is what the publication removed).
    const kept = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: withEntry })
    expect(kept.ok).toBe(true)
    expect(disposition(kept, CAP)).toBe('unverified-nondiscriminating')
    // The proposed publication: exactly scale:scope-disjoint joins the verified Stage 2 set; the comparison ids never do.
    const published = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: withoutEntry })
    expect(published.ok).toBe(true)
    expect(published.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...beforeCap, CAP].sort())
    expect(disposition(published, S2_SCOPE_DISJOINT_TWIN)).toBe('unverified-nondiscriminating')
    expect(disposition(published, S2_SCOPE_DISJOINT_CONTROL)).toBe('unverified-nondiscriminating')
  }, HEAVY)

  for (const h of READINGS.filter(x => x !== 'union')) {
    it(`planted ${h}: discriminating, but not production's reading, so the audit withholds it even without the registry entry`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      expect(CF_STAGE2_PRODUCTION_READINGS[CAP]).toBe('union')
      const { root, manifest: m } = syntheticRoot(h)
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind === 'verified' ? res.matchedHypotheses : res).toEqual({ [CAP]: h })
      const r = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: withoutEntry })
      expect(disposition(r, CAP)).toBe('supported-structural')
      expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(beforeCap)
    }, HEAVY)
  }

  it('a capture no reading explains is a refused capability, never verified (a finding, never fitted)', () => {
    const c = caseOf()
    const base = plantedSidecar(c, 'union') as { observedCells: Array<Record<string, unknown>> }
    // D3 as its per-range twin while every other cell stays union: no reading fits.
    const per = plantedFills(c, 'per-range').get('D3')!
    const ev = judge(c, { observedCells: base.observedCells.map(o => (o.ref === 'D3' ? { ...o, fillRgb: per } : o)) })
    expect([ev[CAP]!.status, ev[CAP]!.code]).toEqual(['refused', 'formula-hypothesis-none'])
  })
})

describe('the reading gate is mandatory from the catalogue, never from the manifest field (codex pre-capture P1)', () => {
  const withoutEntry = CF_STAGE2_NONDISCRIMINATING
  const beforeCap = [...VERIFIED_STAGE2_CAPABILITIES].filter(id => id !== CAP).sort()
  const mutate = (m: CfManifest, fn: (c: CfOracleCase) => CfOracleCase): CfManifest => ({ ...m, cases: m.cases.map(x => (x.id === ID ? fn(x) : x)) })
  const dispositionOf = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }) => r.perCapability.find(p => p.capabilityId === CAP)!.disposition
  const cases: Array<[string, (c: CfOracleCase) => CfOracleCase]> = [
    ['scaleReadings deleted', c => { const { scaleReadings: _r, ...rest } = c; return rest as CfOracleCase }],
    ['scaleReadings empty', c => ({ ...c, scaleReadings: [] })],
    ['scaleReadings edited (per-range relabelled union)', c => ({ ...c, scaleReadings: c.scaleReadings!.map(x => ({ ...x, hypotheses: { ...x.hypotheses, union: x.hypotheses['per-range']!, 'per-range': x.hypotheses.union! } })) })],
  ]
  for (const [name, fn] of cases) {
    it(`synthetic per-range capture, ${name}: the case refuses and scale:scope-disjoint never verifies, even without the registry entry`, async () => {
      const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot('per-range')
      const bad = mutate(m, fn)
      const res = verifyCfCase(bad.cases.find(x => x.id === ID)!, { rootDir: root })
      expect([res.kind, res.kind === 'refusal' ? res.code : undefined]).toEqual(['refusal', 'unobserved-expectation'])
      const r = auditStage2Manifest(bad, { rootDir: root, nonDiscriminating: withoutEntry })
      expect(dispositionOf(r)).not.toBe('verified')
      expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(beforeCap)
    }, HEAVY)
  }

  it('the same mutations on the synthetic union capture refuse too; the unmutated union capture still verifies', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const { root, manifest: m } = syntheticRoot('union')
    for (const [name, fn] of cases) {
      const res = verifyCfCase(mutate(m, fn).cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, name).toBe('refusal')
    }
    expect(dispositionOf(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: withoutEntry }))).toBe('verified')
  }, HEAVY)

  it('a gated capability with no readings evidence in the verified result is withheld by the audit (registry-driven gate)', async () => {
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    // The num-segment follow-up registered nothing (its capture matched no reading).
    expect(Object.keys(CF_STAGE2_PRODUCTION_READINGS)).toEqual([CAP])
    expect(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings!.length).toBeGreaterThan(0)
  })
})

describe('a committed input survives regeneration under either runtime (codex pre-capture P2)', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match, and only then', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    const generated = await generateCaseArchive(ID)
    expect(keepCommittedArchive(committed, generated)).toBe(committed)
    expect(keepCommittedArchive(undefined, generated)).toBe(generated)
    const other = await generateCaseArchive('stage2-scale-scope')
    expect(keepCommittedArchive(committed, other)).toBe(other)
  }, HEAVY)

  it('generateAllStage1 under bun (the build.ts path) leaves the node-generated input, saved file and manifest byte-identical', async () => {
    const { spawnSync } = await import('node:child_process')
    const root = mkdtempSync(join(tmpdir(), 'cf-s2-bun-'))
    const c = caseOf()
    for (const rel of ['tests/fixtures/xlsx-cf/manifest.json', c.input, c.saved]) { mkdirSync(dirname(join(root, rel)), { recursive: true }); copyFileSync(resolve(REPO_ROOT, rel), join(root, rel)) }
    const script = `const { generateAllStage1 } = await import(${JSON.stringify(resolve(REPO_ROOT, 'scripts/cf-refs/generate.ts'))}); await generateAllStage1({ rootDir: ${JSON.stringify(root)}, stage: 2, onlyCaseId: ${JSON.stringify(ID)} })`
    const run = spawnSync('bun', ['-e', script], { cwd: REPO_ROOT, encoding: 'utf8' })
    expect(run.status, run.stdout + run.stderr).toBe(0)
    for (const rel of ['tests/fixtures/xlsx-cf/manifest.json', c.input, c.saved]) expect(sha(new Uint8Array(readFileSync(join(root, rel)))), rel).toBe(sha(new Uint8Array(readFileSync(resolve(REPO_ROOT, rel)))))
  }, HEAVY)
})

describe('the publication: scale:scope-disjoint held', () => {
  it('Stages 1-4: every measured case paints exactly as without scale:scope-disjoint, except this case\'s multi-range cells', async () => {
    const changed: string[] = []
    let compared = 0
    for (const c of manifest.cases.filter(x => x.status === 'measured' && x.stage !== 5)) {
      const refs = c.expected.cells.map(e => e.ref)
      const without = await withExactCfEvidence(WITHOUT_DISJOINT(), () => productionFills(c, refs))
      const held = await productionFills(c, refs)
      for (const ref of refs) { compared++; if (without.get(ref) !== held.get(ref)) changed.push(`${c.id}:${ref}`) }
    }
    const multi = caseOf().scaleReadings!.filter(x => x.role !== 'negative-control').flatMap(x => x.refs).map(ref => `${ID}:${ref}`)
    expect(changed.sort()).toEqual(multi.sort())
    expect(compared).toBeGreaterThan(1000)
  }, HEAVY)

  it('Stage 2: the old stage2-scale-scope disjoint rule stays withheld by its overlaps and hidden row 4', async () => {
    const c = manifest.cases.find(x => x.id === 'stage2-scale-scope')!
    const held = await productionFills(c, ['B2', 'B3', 'D2', 'D3'])
    for (const [ref, fill] of held) expect(fill, ref).toBeUndefined()
  }, HEAVY)

  it('Stage 5: every measured case agrees with Excel with no unjustified refusal; the multi-range scales now paint', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const { stage5JustifiedRefusals } = await import('../scripts/cf-refs/stage5-withhold')
    const published = (id: string): boolean => VERIFIED_CAPABILITIES.has(id) || id === CAP
    const rows = await stage5CaseRecords(manifest, REPO_ROOT, { published })
    expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`)).toEqual([])
    expect(rows.filter(r => r.compared).length).toBeGreaterThan(2000)
    // The multi-range scales of the three cfvo scope cases (real Excel captures) paint, and agree, under union statistics.
    const multi: Array<[string, string]> = [['stage5-cfvo-scope', 'B3'], ['stage5-cfvo-scope', 'D3'], ['stage5-cfvo-scope', 'J3'], ['stage5-cfvo-scope-v2', 'B3'], ['stage5-cfvo-scope-v2', 'K3'], ['stage5-cfvo-relative-a1', 'B3'], ['stage5-cfvo-relative-a1', 'D3'], ['stage5-cfvo-relative-a1', 'H3']]
    for (const [id, ref] of multi) {
      const row = rows.find(r => r.caseId === id && r.ref === ref)!
      expect([row.compared, row.agrees, row.production.unresolved], `${id}:${ref}`).toEqual([true, true, []])
      expect(stage5JustifiedRefusals(manifest.cases.find(x => x.id === id)!, ref, published), `${id}:${ref}`).toEqual([])
    }
  }, 600_000)
})
