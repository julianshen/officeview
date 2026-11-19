// @vitest-environment node
/**
 * Follow-up to the formula:cfvo-scope-multi-range finding (centred glyph key, 2026-11-19; orchestrator ruling 2026-11-20):
 * stage5-cfvo-scope-v2, a discriminating case version.
 *
 * stage5-cfvo-scope holds the relative bound =MAX(B2:B5) on sqref "B2:B5 D2:D5 F2:F5" (a scale, and an icon set on rows
 * 12-15). Excel drew the top colour and the top icon on every cell, contradicting both registered readings (rule-wide,
 * per-range). Every value there is positive and every window outside the data is blank, so a bound of 0, a bound read
 * from an empty window and an ignored bound all paint the same. This case pre-registers twelve readings of the bound:
 *
 * - the window the stored formula names (B2:B7 here), shifted from its frame to an evaluation anchor. The frame is the
 *   sqref's top-left (the documented convention) or A1. The anchor is A1, the sqref's top-left, each range's top-left,
 *   the cell itself or the last range's top-left (frame A1 evaluated at A1 is the written window, i.e. rule-wide);
 * - zero (the bound reads 0), ignored (every value meets it), max-fallback (the bound falls back to the scope's max).
 *
 * Negative data (-40..75) splits zero from ignored, and every window an anchor can read is seeded with known values
 * (spikes and negatives), so each reading predicts its own colour or icon. Twins are literal num/num rules on the same
 * values, one per predicted rendering (canonical: equal renderings share one twin). The case has no capture yet, so its
 * configurations are run through the real Stage 5 evaluator on synthetic observations planted under each reading by an
 * independent model (the Stage 2 sRGB truncate interpolation, pinned below on the stage5-cfvo-scope twins; the Stage 4
 * threshold rule for icons). This checks the family, not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_CASES, CF_STAGE5_GROUPS, CF_STAGE5_NONDISCRIMINATING, CF_STAGE5_OWNERS, CF_STAGE5_UNSUPPORTED, type CfManifest, type CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  evaluateFormulaCapabilities,
  formulaControlsOf,
  readZipPartsSync,
  resolveFormulaConfigurations,
  visualKindOf,
  type FormulaObservation,
  type S5Visual,
} from '../scripts/cf-refs/verify'
import { interpolateScaleColor } from '../src/xlsx/cf/colorScale'

const REPO_ROOT = resolve(__dirname, '..')
const ID = 'stage5-cfvo-scope-v2'
const CAP = 'formula:cfvo-scope-multi-range'
const TWIN = 'formula:cfvo-scope-v2-twin'
const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (id: string): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === id)
  if (!c) throw new Error(`${id} is not in the manifest (generate it with generate.ts --stage 5 --case ${id})`)
  return c
}
const sheetOf = (c: CfOracleCase): { xml: string; parts: Map<string, Uint8Array> } => {
  const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
  return { parts, xml: new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml')!) }
}
const valuesOf = (xml: string): Map<string, number> => {
  const out = new Map<string, number>()
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*><v>([^<]+)<\/v><\/c>/g)) out.set(m[1]!, Number(m[2]))
  return out
}

const colIdx = (l: string): number => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
const colLet = (n: number): string => (n <= 0 ? '' : colLet(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26)))
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
const rows = (col: string, from: number, n: number): string[] => Array.from({ length: n }, (_, i) => `${col}${from + i}`)

const READINGS = [
  'rule-wide', 'per-range', 'anchor-cell', 'anchor-last', 'anchor-a1',
  'a1-frame-top-left', 'a1-frame-per-range', 'a1-frame-cell', 'a1-frame-last',
  'zero', 'ignored', 'max-fallback',
] as const
type Reading = (typeof READINGS)[number]

/**
 * An independent model of each reading's bound: the formula MAX(<first col><r0>:<first col><r0+n-1>), written on the rule's
 * sqref, read under a frame (the cell the stored references are relative to) and an anchor (where Excel evaluates them).
 * A window must be fully seeded unless `blankIsZero` (Excel's MAX over blank cells is 0: the old capture's windows).
 */
function boundsOf(values: Map<string, number>, cols: readonly string[], r0: number, n: number, o: { blankIsZero?: boolean } = {}) {
  const first = colIdx(cols[0]!)
  const last = colIdx(cols.at(-1)!)
  const windows: Record<string, Record<string, string[]>> = {}
  const out: Record<string, Record<string, number>> = {}
  const refs = cols.flatMap(c => rows(c, r0, n))
  const windowAt = (frame: [number, number], anchor: [number, number]): string[] => {
    const top = r0 + anchor[0] - frame[0]
    const col = first + anchor[1] - frame[1]
    return Array.from({ length: n }, (_, i) => `${colLet(col)}${top + i}`)
  }
  const maxOf = (w: string[]): number => {
    const vs = w.map(r => values.get(r))
    if (!o.blankIsZero && vs.some(v => v === undefined)) throw new Error(`window ${w[0]}:${w.at(-1)} is not fully seeded`)
    const nums = vs.filter((v): v is number => v !== undefined)
    return nums.length === 0 ? 0 : Math.max(...nums)
  }
  const TL: [number, number] = [r0, first]
  const A1: [number, number] = [1, 1]
  const anchorOf: Record<string, (ref: string) => [number, number]> = {
    a1: () => A1,
    tl: () => TL,
    range: ref => [r0, colIdx(colOf(ref))],
    cell: ref => [rowOf(ref), colIdx(colOf(ref))],
    last: () => [r0, last],
  }
  const framed: Array<[Reading, [number, number], string]> = [
    ['rule-wide', TL, 'tl'], ['per-range', TL, 'range'], ['anchor-cell', TL, 'cell'], ['anchor-last', TL, 'last'], ['anchor-a1', TL, 'a1'],
    ['a1-frame-top-left', A1, 'tl'], ['a1-frame-per-range', A1, 'range'], ['a1-frame-cell', A1, 'cell'], ['a1-frame-last', A1, 'last'],
  ]
  for (const [h, frame, anchor] of framed) {
    windows[h] = Object.fromEntries(refs.map(r => [r, windowAt(frame, anchorOf[anchor]!(r))]))
    out[h] = Object.fromEntries(refs.map(r => [r, maxOf(windows[h]![r]!)]))
  }
  const scopeMax = Math.max(...refs.map(r => values.get(r)!))
  out.zero = Object.fromEntries(refs.map(r => [r, 0]))
  out.ignored = Object.fromEntries(refs.map(r => [r, -Infinity]))
  out['max-fallback'] = Object.fromEntries(refs.map(r => [r, scopeMax]))
  return { bounds: out as Record<Reading, Record<string, number>>, windows, refs, min: Math.min(...refs.map(r => values.get(r)!)) }
}

/** The Stage 2 colour model: sRGB, truncate; clamped beyond the bounds; inverted bounds swapped (stage2-scale-degenerate D). */
const LO = 0xfff8696b
const HI = 0xff63be7b
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const scaleColour = (v: number, lo: number, hi: number): string => {
  const [a, b] = lo <= hi ? [lo, hi] : [hi, lo]
  if (v >= b) return hex(HI)
  if (v <= a) return hex(LO)
  return hex(interpolateScaleColor(LO, HI, (v - a) / (b - a)))
}
/** The Stage 4 threshold rule: the largest i with v >= t_i (gte), else 0. */
const iconClass = (v: number, t1: number, t2: number): number => (v >= t2 ? 2 : v >= t1 ? 1 : 0)

const SCALE_COLS = ['B', 'D', 'F'] as const
const N = 6
const ICON_R0 = 16
const T1 = -20

describe('stage5-cfvo-scope-v2: the case', () => {
  it('owns formula:cfvo-scope-multi-range (appended last); stage5-cfvo-scope keeps its configurations, capture and finding', () => {
    const c = caseOf(ID)
    expect(c.status).toBe('recipe-pending')
    expect(CF_STAGE5_OWNERS[CAP]).toBe(ID)
    expect(CF_STAGE5_CAPABILITIES.find(e => e.capabilityId === CAP)!.caseId).toBe(ID)
    expect(CF_STAGE5_CASES.at(-1)!.id).toBe(ID)
    expect(CF_STAGE5_GROUPS['stage5-cfvo-scope']).toEqual(['stage5-cfvo-scope', ID])
    const old = CF_STAGE5_CASES.find(x => x.id === 'stage5-cfvo-scope')!
    expect(old.formulaConfigurations!.map(x => x.id)).toEqual(['scope-scale', 'scope-scale-absolute', 'scope-icon'])
    expect(old.status).toBe('recipe-pending')
    expect(caseOf('stage5-cfvo-scope').status).toBe('measured')
    // The finding stays as it is until the capture of this case (task 5); production keeps the bound unresolved.
    expect(CF_STAGE5_UNSUPPORTED[CAP]).toMatch(/no hypothesis \(rule-wide, per-range\) matches/)
    expect(CF_STAGE5_NONDISCRIMINATING[TWIN]).toBeDefined()
  })

  it('the rules: the relative bound on a three-range scale and icon set, an absolute negative control, a single-range control', () => {
    const c = caseOf(ID)
    const legacy = c.markers.filter(m => m.capabilityId !== TWIN)
    expect(legacy.map(m => [m.capabilityId, m.ruleType, m.sqref])).toEqual([
      [CAP, 'colorScale', 'B2:B7 D2:D7 F2:F7'],
      [CAP, 'iconSet', 'B16:B21 D16:D21 F16:F21'],
      [CAP, 'colorScale', 'K2:K7 M2:M7 O2:O7'],
      ['formula:cfvo-relative', 'colorScale', 'I2:I7'],
    ])
    expect(legacy[0]!.scale!.cfvo).toEqual([{ type: 'min' }, { type: 'formula', f: 'MAX(B2:B7)' }])
    expect(legacy[1]!.icon!.cfvo).toEqual([{ type: 'percent', val: '0' }, { type: 'num', val: String(T1) }, { type: 'formula', val: 'MAX(B16:B21)' }])
    expect(legacy[2]!.scale!.cfvo).toEqual([{ type: 'min' }, { type: 'formula', f: '$H$1' }])
    expect(legacy[3]!.scale!.cfvo).toEqual([{ type: 'min' }, { type: 'formula', f: 'MAX(I2:I7)' }])
    // Every twin is a literal num/num rule (a scale) or a num/num icon set, never a formula.
    for (const m of c.markers.filter(x => x.capabilityId === TWIN)) {
      const cfvo = m.scale?.cfvo ?? m.icon!.cfvo.slice(1)
      expect(cfvo.every(x => x.type === 'num'), m.sqref).toBe(true)
    }
    const cfgs = c.formulaConfigurations!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.observable, x.role ?? x.record ?? '', x.replicateBy ?? ''])).toEqual([
      ['scope2-scale', CAP, 'fill', '', 'column'],
      ['scope2-icon', CAP, 'crop', '', 'column'],
      ['scope2-scale-absolute', CAP, 'fill', 'negative-control', ''],
      ['scope2-single-range', 'formula:cfvo-relative', 'fill', expect.stringMatching(/^Recorded only/), ''],
    ])
    for (const cfg of cfgs) expect(Object.keys(cfg.hypotheses), cfg.id).toEqual([...READINGS])
  })

  it('20 pt rows and the default (vertically centred) style, the frame of the centred glyph key', () => {
    const { xml } = sheetOf(caseOf(ID))
    expect(/<sheetFormatPr defaultRowHeight="20"/.test(xml)).toBe(true)
    expect(xml).not.toMatch(/ s="\d+"/)
    // Every value is at most three characters, so the centred number stays clear of the 28 px icon crop.
    for (const v of valuesOf(xml).values()) expect(String(v).length, String(v)).toBeLessThanOrEqual(3)
  })

  it('the data: negatives in every range, one shared minimum per rule, and the windows every reading reads are fully seeded', () => {
    const c = caseOf(ID)
    const values = valuesOf(sheetOf(c).xml)
    const scale = boundsOf(values, SCALE_COLS, 2, N)
    const icon = boundsOf(values, SCALE_COLS, ICON_R0, N)
    const single = boundsOf(values, ['I'], 2, N)
    for (const col of SCALE_COLS) {
      // 'min' is the same for the union and for each range, so a per-range minimum cannot confound the bound.
      expect(Math.min(...rows(col, 2, N).map(r => values.get(r)!)), col).toBe(scale.min)
      expect(rows(col, 2, N).filter(r => values.get(r)! < 0).length, col).toBeGreaterThanOrEqual(2)
      expect(rows(col, ICON_R0, N).filter(r => values.get(r)! < 0).length, col).toBeGreaterThanOrEqual(2)
    }
    // Non-inverted: every finite scale bound lies above the minimum; every finite icon bound at or above the num threshold.
    for (const b of [scale, single]) for (const h of READINGS) for (const [r, v] of Object.entries(b.bounds[h])) if (Number.isFinite(v)) expect(v, `${h} ${r}`).toBeGreaterThan(b.min)
    for (const h of READINGS) for (const [r, v] of Object.entries(icon.bounds[h])) if (Number.isFinite(v)) expect(v, `${h} ${r}`).toBeGreaterThanOrEqual(T1)
    // No window touches a twin, a control or another configuration's cells; it reads its own rule's data or its seeds.
    const designated = new Set(c.expected.cells.map(e => e.ref))
    const sqrefCells = (sq: string): Set<string> => new Set(sq.split(' ').flatMap(t => { const [a, b] = t.split(':'); return rows(colOf(a!), rowOf(a!), rowOf(b ?? a!) - rowOf(a!) + 1) }))
    const rules = c.markers.map(m => ({ sqref: m.sqref, cells: sqrefCells(m.sqref) }))
    for (const [b, own] of [[scale, 'B2:B7 D2:D7 F2:F7'], [icon, 'B16:B21 D16:D21 F16:F21'], [single, 'I2:I7']] as const) {
      for (const [h, perRef] of Object.entries(b.windows)) for (const w of Object.values(perRef)) for (const cell of w) {
        const inOwn = sqrefCells(own).has(cell)
        expect(inOwn || !rules.some(r => r.cells.has(cell)), `${h}: ${cell} lies in another rule`).toBe(true)
        expect(inOwn || !designated.has(cell), `${h}: ${cell} is designated`).toBe(true)
      }
    }
  })

  it('the colour model is the one Excel drew on the stage5-cfvo-scope twins (sRGB, truncate)', () => {
    const old = caseOf('stage5-cfvo-scope')
    const side = JSON.parse(readFileSync(resolve(REPO_ROOT, old.captures[0]!.sidecarPath), 'utf8')) as { formulaObservations: Record<string, { displayFormatFill: string }> }
    const fill = (r: string): string => side.formulaObservations[r]!.displayFormatFill
    // P: 10/20/30/40 under num 10/40; S: 10/20/30/60 under num 10/60; J/U: 10/20/30/40 under 10/50.
    expect([fill('P2'), fill('P3'), fill('P4'), fill('P5')]).toEqual([10, 20, 30, 40].map(v => scaleColour(v, 10, 40)))
    expect([fill('P3'), fill('P4')]).toEqual(['FFC78570', 'FF95A175'])
    expect([fill('S3'), fill('S4'), fill('U3'), fill('U5')]).toEqual([scaleColour(20, 10, 60), scaleColour(30, 10, 60), scaleColour(20, 10, 50), scaleColour(40, 10, 50)])
  })
})

/** Planted observations: each designated cell under one reading, each twin under its own literal rule, controls blank. */
function plant(c: CfOracleCase, values: Map<string, number>, reading: Reading): FormulaObservation {
  const scale = boundsOf(values, SCALE_COLS, 2, N)
  const icon = boundsOf(values, SCALE_COLS, ICON_R0, N)
  const single = boundsOf(values, ['I'], 2, N)
  const absolute = values.get('H1')!
  const absRefs = new Set(['K', 'M', 'O'].flatMap(col => rows(col, 2, N)))
  const twinOf = (ref: string) => c.markers.find(m => m.capabilityId === TWIN && m.sqref.split(':')[0]!.replace(/\d+/, '') === colOf(ref) && rowOf(ref) >= rowOf(m.sqref.split(':')[0]!) && rowOf(ref) <= rowOf(m.sqref.split(':')[1]!))
  const num = (x: { val?: string; f?: string }): number => Number(x.val ?? x.f)
  return {
    fill(ref) {
      const v = values.get(ref)
      const none = { fill: 'none', strike: false }
      if (v === undefined) return none
      if (scale.bounds[reading][ref] !== undefined) return { fill: scaleColour(v, scale.min, scale.bounds[reading][ref]!), strike: false }
      if (single.bounds[reading][ref] !== undefined) return { fill: scaleColour(v, single.min, single.bounds[reading][ref]!), strike: false }
      if (absRefs.has(ref)) return { fill: scaleColour(v, scale.min, absolute), strike: false }
      const t = twinOf(ref)
      if (t?.scale) return { fill: scaleColour(v, num(t.scale.cfvo[0]!), num(t.scale.cfvo[1]!)), strike: false }
      return none
    },
    visual(ref) {
      const v = values.get(ref)
      const at = (k: number): S5Visual => ({ kind: 'icon', cls: `3Arrows:${k}`, left: 2, top: 4 })
      if (v === undefined) return { kind: 'none' }
      if (icon.bounds[reading][ref] !== undefined) return at(iconClass(v, T1, icon.bounds[reading][ref]!))
      const t = twinOf(ref)
      if (t?.icon) return at(iconClass(v, num(t.icon.cfvo[1]!), num(t.icon.cfvo[2]!)))
      return { kind: 'none' }
    },
  }
}

describe('stage5-cfvo-scope-v2: every reading discriminates on a synthetic observation', () => {
  for (const reading of READINGS) {
    it(`planted ${reading}: the capability comes back discriminating, matched ${reading}`, () => {
      const c = caseOf(ID)
      const { xml, parts } = sheetOf(c)
      const ev = evaluateFormulaCapabilities(resolveFormulaConfigurations(c, parts), formulaControlsOf(c), plant(c, valuesOf(xml), reading), { kindOf: visualKindOf(c) })[CAP]!
      expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', reading])
    }, 60_000)
  }

  it('every pair of readings differs at >= 2 cells of every range where it differs at all (colour), and at >= 2 cells overall', () => {
    const values = valuesOf(sheetOf(caseOf(ID)).xml)
    const scale = boundsOf(values, SCALE_COLS, 2, N)
    const icon = boundsOf(values, SCALE_COLS, ICON_R0, N)
    const colour = (h: Reading, r: string): string => scaleColour(values.get(r)!, scale.min, scale.bounds[h][r]!)
    const cls = (h: Reading, r: string): number => iconClass(values.get(r)!, T1, icon.bounds[h][r]!)
    for (const [i, a] of READINGS.entries()) for (const b of READINGS.slice(i + 1)) {
      let total = 0
      for (const col of SCALE_COLS) {
        const s = rows(col, 2, N).filter(r => colour(a, r) !== colour(b, r)).length
        const k = rows(col, ICON_R0, N).filter(r => cls(a, r) !== cls(b, r)).length
        expect(s === 0 || s >= 2, `${a}/${b} scale ${col}: ${s}`).toBe(true)
        expect(k === 0 || k >= 2, `${a}/${b} icon ${col}: ${k}`).toBe(true)
        total += s + k
      }
      expect(total, `${a}/${b}`).toBeGreaterThanOrEqual(2)
    }
  })

  it('the single-range control separates the readings that stay distinct on one range (recorded only, never evidence)', () => {
    const values = valuesOf(sheetOf(caseOf(ID)).xml)
    const single = boundsOf(values, ['I'], 2, N)
    const colour = (h: Reading, r: string): string => scaleColour(values.get(r)!, single.min, single.bounds[h][r]!)
    const key = (h: Reading): string => rows('I', 2, N).map(r => colour(h, r)).join()
    // On one range: per-range, anchor-last and max-fallback are rule-wide; the a1-frame anchors other than the cell coincide.
    for (const h of ['per-range', 'anchor-last', 'max-fallback'] as const) expect(key(h), h).toBe(key('rule-wide'))
    for (const h of ['a1-frame-per-range', 'a1-frame-last'] as const) expect(key(h), h).toBe(key('a1-frame-top-left'))
    const distinct = ['rule-wide', 'anchor-cell', 'anchor-a1', 'a1-frame-top-left', 'a1-frame-cell', 'zero', 'ignored'] as const
    for (const [i, a] of distinct.entries()) for (const b of distinct.slice(i + 1)) {
      expect(rows('I', 2, N).filter(r => colour(a, r) !== colour(b, r)).length, `${a}/${b}`).toBeGreaterThanOrEqual(2)
    }
  })
})

describe('retrodiction: what the two committed captures already say about each reading', () => {
  it('stage5-cfvo-scope (all top) contradicts the readings whose window holds its data; a zero, blank-window or ignored bound fits', () => {
    const old = caseOf('stage5-cfvo-scope')
    const values = valuesOf(sheetOf(old).xml)
    const side = JSON.parse(readFileSync(resolve(REPO_ROOT, old.captures[0]!.sidecarPath), 'utf8')) as { formulaObservations: Record<string, { displayFormatFill: string }> }
    const b = boundsOf(values, SCALE_COLS, 2, 4, { blankIsZero: true })
    const fits = READINGS.filter(h => b.refs.every(r => scaleColour(values.get(r)!, b.min, b.bounds[h][r]!) === side.formulaObservations[r]!.displayFormatFill))
    expect(b.refs.every(r => side.formulaObservations[r]!.displayFormatFill === hex(HI))).toBe(true)
    expect(fits).toEqual(['anchor-a1', 'a1-frame-top-left', 'a1-frame-per-range', 'a1-frame-cell', 'a1-frame-last', 'zero', 'ignored'])
  })

  it('stage5-cfvo-formulas K (the single-range relative lower bound =A2, matched zero) contradicts every on-sheet relative reading', () => {
    const old = caseOf('stage5-cfvo-formulas')
    const values = valuesOf(sheetOf(old).xml)
    const side = JSON.parse(readFileSync(resolve(REPO_ROOT, old.captures[0]!.sidecarPath), 'utf8')) as { formulaObservations: Record<string, { displayFormatFill: string }> }
    const K = rows('K', 2, 8)
    const max = Math.max(...K.map(r => values.get(r)!))
    // The lower bound under each reading (=A2 written on K2:K9): its frame K2 (or A1), evaluated at K2 or at the cell.
    const lower: Record<string, (r: string) => number> = {
      'rule-wide': () => values.get('A2')!,
      'anchor-cell': r => values.get(`A${rowOf(r)}`)!,
      'a1-frame-top-left': () => values.get('K3')!,
      'a1-frame-cell': r => values.get(`K${rowOf(r) + 1}`) ?? 0,
      zero: () => 0,
    }
    const fits = Object.keys(lower).filter(h => K.every(r => scaleColour(values.get(r)!, lower[h]!(r), max) === side.formulaObservations[r]!.displayFormatFill))
    expect(fits).toEqual(['zero'])
    // anchor-a1 there reads ten columns left of A1 (off the sheet: a wrap or #REF!), which the case cannot model.
  })
})
