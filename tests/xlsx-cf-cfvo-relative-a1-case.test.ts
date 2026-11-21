// @vitest-environment node
/**
 * The cfvo-relative confirmation case (brief 2026-11-21, Phase 1 Task 2): stage5-cfvo-relative-a1.
 *
 * stage5-cfvo-scope-v2 (5dc2fe9e) matched `anchor-a1` (a relative cfvo bound read as if its formula sat in A1) on several
 * ranges and, in its recorded single-range control, on one. Production's anchor-a1 post-dates that capture, so this case,
 * built and captured after it, pre-registers anchor-a1 and its rivals for both formula:cfvo-scope-multi-range and
 * formula:cfvo-relative, on a scale, an icon set and a data bar each. It has no capture yet: every reading is planted on a
 * synthetic observation by an independent model (the Stage 2 sRGB truncate scale, the Stage 4 threshold rule, the 84 px
 * Stage 5 bar of tests/xlsx-cf-bar-bounds-formula.test.ts) and run through the real Stage 5 evaluator. This checks the
 * family, not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_CASES, CF_STAGE5_GROUPS, CF_STAGE5_NONDISCRIMINATING, CF_STAGE5_OWNERS, CF_STAGE5_UNSUPPORTED, type CfManifest, type CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { evaluateFormulaCapabilities, formulaControlsOf, readZipPartsSync, resolveFormulaConfigurations, visualKindOf, type FormulaObservation, type S5Visual } from '../scripts/cf-refs/verify'
import { interpolateScaleColor } from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, injectCfFormulaSelections, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { injectCfFormulaReadings } from '../src/xlsx/cf/formulaReadings'
import { VERIFIED_STAGE5_FORMULA_SELECTIONS } from '../src/xlsx/cf/verifiedCapabilities'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { caseEvaluationOf, caseNowOf, caseParseOptionsOf } from '../scripts/cf-refs/case-clock'

const REPO_ROOT = resolve(__dirname, '..')
const ID = 'stage5-cfvo-relative-a1'
const MULTI = 'formula:cfvo-scope-multi-range'
const SINGLE = 'formula:cfvo-relative'
const TWIN = 'formula:cfvo-relative-a1-twin'
const N = 4
const T1 = -20
const MULTI_READINGS = ['anchor-a1', 'rule-wide', 'per-range', 'anchor-cell', 'zero', 'ignored'] as const
const SINGLE_READINGS = ['anchor-a1', 'anchor-top-left', 'anchor-cell', 'zero', 'ignored'] as const
/** Bars carry the window readings only: zero and ignored would invert them (their twins refused by production, S3-E). */
const windowsOnly = <T extends string>(rs: readonly T[]): T[] => rs.filter(h => h !== 'zero' && h !== 'ignored')

const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === ID)
  if (!c) throw new Error(`${ID} is not in the manifest (generate it with generate.ts --stage 5 --case ${ID})`)
  return c
}
const partsOf = (c: CfOracleCase, file = c.saved): Map<string, Uint8Array> => readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, file))))!
const valuesOf = (c: CfOracleCase): Map<string, number> => {
  const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
  const out = new Map<string, number>()
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*><v>([^<]+)<\/v><\/c>/g)) out.set(m[1]!, Number(m[2]))
  return out
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
const rows = (col: string, from: number): string[] => Array.from({ length: N }, (_, i) => `${col}${from + i}`)

/** An independent model of each reading's bound for `=MAX(<cols[0]><r0>:...)` written on the rule's top-left. */
function boundsOf(values: Map<string, number>, cols: readonly string[], r0: number, readings: readonly string[]) {
  const windows: Record<string, Record<string, string[]>> = {}
  const bounds: Record<string, Record<string, number>> = {}
  const win = (col: string, row: number): string[] => Array.from({ length: N }, (_, i) => `${col}${row + i}`)
  const of: Record<string, (ref: string) => string[] | undefined> = {
    'anchor-a1': () => win('A', 1), 'rule-wide': () => win(cols[0]!, r0), 'anchor-top-left': () => win(cols[0]!, r0),
    'per-range': ref => win(colOf(ref), r0), 'anchor-cell': ref => win(colOf(ref), rowOf(ref)), zero: () => undefined, ignored: () => undefined,
  }
  const refs = cols.flatMap(c => rows(c, r0))
  for (const h of readings) {
    windows[h] = {}
    bounds[h] = {}
    for (const r of refs) {
      const w = of[h]!(r)
      if (w) {
        const vs = w.map(x => values.get(x))
        if (vs.some(v => v === undefined)) throw new Error(`${h} ${r}: window ${w[0]}:${w.at(-1)} is not fully seeded`)
        windows[h]![r] = w
        bounds[h]![r] = Math.max(...(vs as number[]))
      } else bounds[h]![r] = h === 'zero' ? 0 : -Infinity
    }
  }
  return { bounds, windows, refs, min: Math.min(...refs.map(r => values.get(r)!)) }
}
const LO = 0xfff8696b, HI = 0xff63be7b
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const scaleColour = (v: number, lo: number, hi: number): string => (v >= hi ? hex(HI) : v <= lo ? hex(LO) : hex(interpolateScaleColor(LO, HI, (v - lo) / (hi - lo))))
const iconClass = (v: number, t1: number, t2: number): number => (v >= t2 ? 2 : v >= t1 ? 1 : 0)
/**
 * Excel's solid 0/100 bar in a Stage 5 cell (84 px: 2 px inset, 81 px span; tests/xlsx-cf-bar-bounds-formula.test.ts),
 * clamped beyond its bounds and drawn with inverted bounds swapped (stage3-bar-degenerate E and D).
 */
const barOf = (v: number, lo: number, hi: number): S5Visual => {
  const [a, b] = lo <= hi ? [lo, hi] : [hi, lo]
  if (v >= b) return { kind: 'bar', left: 2, right: 82 }
  const f = (v - a) / (b - a)
  return f <= 0 ? { kind: 'none' } : { kind: 'bar', left: 2, right: 1 + Math.round(f * 81) }
}

const BLOCKS = [
  { cap: MULTI, cfg: 'rel-a1-scale', kind: 'scale', cols: ['B', 'D'], r0: 2, readings: MULTI_READINGS },
  { cap: MULTI, cfg: 'rel-a1-icon', kind: 'icon', cols: ['B', 'D'], r0: 11, readings: MULTI_READINGS },
  { cap: MULTI, cfg: 'rel-a1-bar', kind: 'bar', cols: ['B', 'D'], r0: 20, readings: windowsOnly(MULTI_READINGS) },
  { cap: SINGLE, cfg: 'rel-a1-single-scale', kind: 'scale', cols: ['F'], r0: 2, readings: SINGLE_READINGS },
  { cap: SINGLE, cfg: 'rel-a1-single-icon', kind: 'icon', cols: ['F'], r0: 11, readings: SINGLE_READINGS },
  { cap: SINGLE, cfg: 'rel-a1-single-bar', kind: 'bar', cols: ['F'], r0: 20, readings: windowsOnly(SINGLE_READINGS) },
] as const

describe('stage5-cfvo-relative-a1: the case', () => {
  it('owns both cfvo scope ids (appended last); the older versions keep their configurations and captures', () => {
    const c = caseOf()
    expect(c.status).toBe('recipe-pending')
    for (const cap of [MULTI, SINGLE]) {
      expect(CF_STAGE5_OWNERS[cap], cap).toBe(ID)
      expect(CF_STAGE5_CAPABILITIES.find(e => e.capabilityId === cap)!.caseId, cap).toBe(ID)
      // Withheld until this case is captured and a reading registered (nothing is registered or published here).
      expect(CF_STAGE5_UNSUPPORTED[cap], cap).toBeDefined()
    }
    expect(CF_STAGE5_CASES.at(-1)!.id).toBe(ID)
    expect(CF_STAGE5_GROUPS['stage5-cfvo-scope']).toContain(ID)
    expect(CF_STAGE5_NONDISCRIMINATING[TWIN]).toBeDefined()
    for (const old of ['stage5-cfvo-scope', 'stage5-cfvo-scope-v2', 'stage5-cfvo-formulas']) expect(manifest.cases.find(x => x.id === old)!.status, old).toBe('measured')
  })

  it('the rules: a scale, an icon set and a data bar on two ranges and on one, plus an absolute negative control', () => {
    const c = caseOf()
    const own = c.markers.filter(m => m.capabilityId !== TWIN && m.extension === 'legacy')
    expect(own.map(m => [m.capabilityId, m.ruleType, m.sqref])).toEqual([
      [MULTI, 'colorScale', 'B2:B5 D2:D5'], [MULTI, 'iconSet', 'B11:B14 D11:D14'], [MULTI, 'dataBar', 'B20:B23 D20:D23'], [MULTI, 'colorScale', 'H2:H5 J2:J5'],
      [SINGLE, 'colorScale', 'F2:F5'], [SINGLE, 'iconSet', 'F11:F14'], [SINGLE, 'dataBar', 'F20:F23'],
    ])
    const cfgs = c.formulaConfigurations!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.observable, x.role ?? '', x.replicateBy ?? ''])).toEqual([
      ['rel-a1-scale', MULTI, 'fill', '', 'column'], ['rel-a1-icon', MULTI, 'crop', '', 'column'], ['rel-a1-bar', MULTI, 'crop', '', 'column'],
      ['rel-a1-scale-absolute', MULTI, 'fill', 'negative-control', ''],
      ['rel-a1-single-scale', SINGLE, 'fill', '', ''], ['rel-a1-single-icon', SINGLE, 'crop', '', ''], ['rel-a1-single-bar', SINGLE, 'crop', '', ''],
    ])
    for (const b of BLOCKS) expect(Object.keys(cfgs.find(x => x.id === b.cfg)!.hypotheses), b.cfg).toEqual([...b.readings])
    // The bars omit zero and ignored, with the reason; no twin of any configuration lies beyond its own bounds (S3-E).
    for (const id of ['rel-a1-bar', 'rel-a1-single-bar']) expect(Object.keys(cfgs.find(x => x.id === id)!.omitted ?? {}), id).toEqual(['zero', 'ignored'])
    const values = valuesOf(c)
    for (const m of c.markers.filter(x => x.capabilityId === TWIN && x.ruleType === 'dataBar' && x.extension === 'legacy')) {
      const [lo, hi] = m.bar!.cfvo.map(x => Number(x.val))
      for (const r of rows(colOf(m.sqref), rowOf(m.sqref.split(':')[0]!))) expect(values.get(r)! >= lo! && values.get(r)! <= hi!, `${m.sqref} ${r}`).toBe(true)
    }
    for (const m of c.markers.filter(x => x.capabilityId === TWIN)) {
      const cfvo = m.scale?.cfvo ?? m.bar?.cfvo ?? m.icon!.cfvo.slice(1)
      expect(cfvo.every(x => x.type === 'num'), m.sqref).toBe(true)
    }
  })

  it('smaller than stage5-cfvo-scope-v2: 20 pt rows, the default centred style, short values', () => {
    const c = caseOf()
    expect(c.expected.cells.length).toBeLessThan(manifest.cases.find(x => x.id === 'stage5-cfvo-scope-v2')!.expected.cells.length)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(/<sheetFormatPr defaultRowHeight="20"/.test(xml)).toBe(true)
    expect(xml).not.toMatch(/ s="\d+"/)
    for (const v of valuesOf(c).values()) expect(String(v).length, String(v)).toBeLessThanOrEqual(3)
  })

  it('every window is fully seeded and reads its own data or its seeds; no bound is inverted or below the icon threshold', () => {
    const c = caseOf()
    const values = valuesOf(c)
    const designated = new Set(c.expected.cells.map(e => e.ref))
    for (const b of BLOCKS) {
      const { bounds, windows, min, refs } = boundsOf(values, b.cols, b.r0, b.readings)
      const own = new Set(refs)
      for (const [h, per] of Object.entries(windows)) for (const w of Object.values(per)) for (const cell of w) {
        expect(own.has(cell) || !designated.has(cell), `${b.cfg} ${h}: ${cell} is designated`).toBe(true)
        expect(own.has(cell) || !c.markers.some(m => m.sqref.split(' ').some(t => { const [a, z] = t.split(':'); return colOf(a!) === colOf(cell) && rowOf(cell) >= rowOf(a!) && rowOf(cell) <= rowOf(z ?? a!) })), `${b.cfg} ${h}: ${cell} lies in a rule`).toBe(true)
      }
      for (const h of b.readings) for (const [r, v] of Object.entries(bounds[h]!)) {
        if (!Number.isFinite(v)) continue
        if (b.kind === 'icon') expect(v, `${b.cfg} ${h} ${r}`).toBeGreaterThanOrEqual(T1)
        else expect(v, `${b.cfg} ${h} ${r}`).toBeGreaterThan(min)
        // S3-E: no bar value beyond a reading's bound (production would refuse the whole rule).
        if (b.kind === 'bar') expect(values.get(r)!, `${b.cfg} ${h} ${r}`).toBeLessThanOrEqual(v)
      }
    }
  })

  it('every bar edge is clear of the value text (<= 61 px) or full width, and off half pixels', () => {
    const values = valuesOf(caseOf())
    for (const b of BLOCKS.filter(x => x.kind === 'bar')) {
      const { bounds, min, refs } = boundsOf(values, b.cols, b.r0, b.readings)
      for (const h of b.readings) for (const r of refs) {
        const bar = barOf(values.get(r)!, min, bounds[h]![r]!)
        if (bar.kind === 'bar') expect(bar.right <= 61 || bar.right === 82, `${b.cfg} ${h} ${r}: ${bar.right}`).toBe(true)
        const x = ((values.get(r)! - min) / (bounds[h]![r]! - min)) * 81
        if (x > 0 && x < 81) expect(Math.abs((x % 1) - 0.5) >= 0.15, `${b.cfg} ${h} ${r}: ${x}`).toBe(true)
      }
    }
  })
})

/** Planted observations: designated cells under one reading per capability, twins under their own literal rule. */
function plant(c: CfOracleCase, values: Map<string, number>, reading: Readonly<Record<string, string>>): FormulaObservation {
  const twinOf = (ref: string) => c.markers.find(m => m.capabilityId === TWIN && m.extension === 'legacy' && colOf(m.sqref) === colOf(ref) && rowOf(ref) >= rowOf(m.sqref.split(':')[0]!) && rowOf(ref) <= rowOf(m.sqref.split(':')[1]!))
  const n = (x: { val?: string; f?: string }): number => Number(x.val ?? x.f)
  const blockOf = (ref: string) => BLOCKS.find(b => (b.cols as readonly string[]).includes(colOf(ref)) && rowOf(ref) >= b.r0 && rowOf(ref) < b.r0 + N)
  const abs = new Set(['H', 'J'].flatMap(col => rows(col, 2)))
  const absMin = Math.min(...[...abs].map(r => values.get(r)!))
  const designated = (ref: string): { kind: string; v: number; lo: number; b: number } | undefined => {
    const blk = blockOf(ref)
    const v = values.get(ref)
    if (!blk || v === undefined) return undefined
    // A reading the configuration omits (zero, ignored on a bar) still renders in the planted world: an inverted bar,
    // drawn swapped (stage3-bar-degenerate D), full.
    const h = reading[blk.cap]!
    const { bounds, min } = boundsOf(values, blk.cols, blk.r0, [h])
    return { kind: blk.kind, v, lo: min, b: bounds[h]![ref]! }
  }
  return {
    fill(ref) {
      const v = values.get(ref)
      const none = { fill: 'none', strike: false }
      if (v === undefined) return none
      if (abs.has(ref)) return { fill: scaleColour(v, absMin, values.get('H1')!), strike: false }
      const d = designated(ref)
      if (d) return d.kind === 'scale' ? { fill: scaleColour(d.v, d.lo, d.b), strike: false } : none
      const t = twinOf(ref)
      if (t?.scale) return { fill: scaleColour(v, n(t.scale.cfvo[0]!), n(t.scale.cfvo[1]!)), strike: false }
      return none
    },
    visual(ref) {
      const v = values.get(ref)
      if (v === undefined) return { kind: 'none' }
      const d = designated(ref)
      if (d?.kind === 'icon') return { kind: 'icon', cls: `3Arrows:${iconClass(d.v, T1, d.b)}`, left: 2, top: 4 }
      if (d?.kind === 'bar') return barOf(d.v, d.lo, d.b)
      const t = twinOf(ref)
      if (t?.icon) return { kind: 'icon', cls: `3Arrows:${iconClass(v, n(t.icon.cfvo[1]!), n(t.icon.cfvo[2]!))}`, left: 2, top: 4 }
      if (t?.bar) return barOf(v, n(t.bar.cfvo[0]!), n(t.bar.cfvo[1]!))
      return { kind: 'none' }
    },
  }
}

describe('stage5-cfvo-relative-a1: each reading of each capability discriminates on a synthetic observation', () => {
  const judge = (reading: Readonly<Record<string, string>>) => {
    const c = caseOf()
    return evaluateFormulaCapabilities(resolveFormulaConfigurations(c, partsOf(c)), formulaControlsOf(c), plant(c, valuesOf(c), reading), { kindOf: visualKindOf(c) })
  }
  for (const h of MULTI_READINGS) {
    it(`planted ${MULTI} = ${h} (and ${SINGLE} = anchor-a1): ${h} comes back as the one match`, () => {
      const ev = judge({ [MULTI]: h, [SINGLE]: 'anchor-a1' })
      expect([ev[MULTI]!.status, ev[MULTI]!.matched], ev[MULTI]!.reason).toEqual(['discriminating', h])
    }, 300_000)
  }
  for (const h of SINGLE_READINGS) {
    it(`planted ${SINGLE} = ${h} (and ${MULTI} = anchor-a1): ${h} comes back as the one match`, () => {
      const ev = judge({ [MULTI]: 'anchor-a1', [SINGLE]: h })
      expect([ev[SINGLE]!.status, ev[SINGLE]!.matched], ev[SINGLE]!.reason).toEqual(['discriminating', h])
    }, 300_000)
  }

  it('distinct predicted colours stay distinct if Excel computed any channel 0.1 of a unit off', () => {
    const values = valuesOf(caseOf())
    const channels = (v: number, lo: number, b: number): Array<[number, number]> => {
      const ends = [[0xf8, 0x63], [0x69, 0xbe], [0x6b, 0x7b]] as const
      if (v >= b) return ends.map(([, x]) => [x, x])
      if (v <= lo) return ends.map(([x]) => [x, x])
      const f = (v - lo) / (b - lo)
      return ends.map(([c0, c1]) => { const d = Math.abs(c1 - c0), s = c1 >= c0 ? 1 : -1, a = c0 + s * Math.floor(f * d - 0.1), z = c0 + s * Math.floor(f * d + 0.1); return [Math.min(a, z), Math.max(a, z)] })
    }
    const apart = (x: Array<[number, number]>, y: Array<[number, number]>): boolean => x.some((p, i) => p[1] < y[i]![0] || y[i]![1] < p[0])
    for (const b of BLOCKS.filter(x => x.kind === 'scale')) {
      const { bounds, min, refs } = boundsOf(values, b.cols, b.r0, b.readings)
      for (const r of refs) {
        const seen = new Map<string, Array<[number, number]>>()
        for (const h of b.readings) seen.set(scaleColour(values.get(r)!, min, bounds[h]![r]!), channels(values.get(r)!, min, bounds[h]![r]!))
        const list = [...seen]
        for (const [i, [ka, ca]] of list.entries()) for (const [kb, cb] of list.slice(i + 1)) expect(apart(ca, cb), `${b.cfg} ${r}: ${ka} / ${kb}`).toBe(true)
      }
    }
  })
})

/**
 * Pre-capture fix round 1 (codex P1b): a rehearsal of the proposed registration. With both formula capabilities held and
 * read as anchor-a1 (the test seams: a proposed registry entry and a selection), and every other capability as published
 * today, the production counterfactual over this case, on the planted anchor-a1 observation, must hold no unjustified
 * refusal and agree with that observation at every designated and control cell. scale:scope-disjoint stays withheld.
 */
describe('stage5-cfvo-relative-a1: rehearsing the proposed anchor-a1 registration against production', () => {
  const PROPOSED = { [SINGLE]: 'anchor-a1', [MULTI]: ['rule-wide', 'per-range', 'anchor-a1'] } as const
  it('the multi-range scales are refused by scale:scope-disjoint alone, a recorded justification; the icon sets and bars are not', async () => {
    const { stage5JustifiedRefusals } = await import('../scripts/cf-refs/stage5-withhold')
    const c = caseOf()
    const held = new Set([...VERIFIED_CAPABILITIES, MULTI, SINGLE])
    expect(held.has('scale:scope-disjoint')).toBe(false)
    injectCfFormulaReadings(PROPOSED)
    injectCfFormulaSelections({ ...VERIFIED_STAGE5_FORMULA_SELECTIONS, [MULTI]: 'anchor-a1' })
    try {
      const reasons = (ref: string): string[] => stage5JustifiedRefusals(c, ref, id => held.has(id)).map(r => `${r.rule}: ${r.reason}`)
      expect(reasons('B2')).toEqual([`${MULTI}/colorScale@p1: scale:scope-disjoint not held (a visual over more than one range)`])
      expect(reasons('H2')).toEqual([`${MULTI}/colorScale@p4: scale:scope-disjoint not held (a visual over more than one range)`])
      for (const ref of ['B11', 'D20', 'F2', 'F11', 'F20']) expect(reasons(ref), ref).toEqual([])
      // With scale:scope-disjoint held as well, nothing on this case is a justified refusal.
      expect(stage5JustifiedRefusals(c, 'B2', id => held.has(id) || id === 'scale:scope-disjoint')).toEqual([])
    } finally { injectCfFormulaReadings(null); injectCfFormulaSelections(null) }
  })

  it('the production counterfactual over the planted anchor-a1 world: no unjustified refusal, no disagreement, and the icon sets, bars and single-range scale paint', async () => {
    const { stage5CounterfactualOf, stage5CellAgreement, stage5TokenOf } = await import('../scripts/cf-refs/stage5-withhold')
    const c = caseOf()
    const values = valuesOf(c)
    const held = [...VERIFIED_CAPABILITIES, MULTI, SINGLE]
    const obs = plant(c, values, { [MULTI]: 'anchor-a1', [SINGLE]: 'anchor-a1' })
    injectCfFormulaReadings(PROPOSED)
    injectCfFormulaSelections({ ...VERIFIED_STAGE5_FORMULA_SELECTIONS, [MULTI]: 'anchor-a1' })
    try {
      await withInjectedCfEvidence(held, async () => {
        const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))), caseParseOptionsOf(c))
        const sheet = doc.sheets[0]!
        const refs = c.expected.cells.map(e => e.ref)
        const truths = await stage5CounterfactualOf({ c, doc, sheet, refs, now: caseNowOf(c), published: id => held.includes(id) })
        const memory = createCfMemory()
        const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation: caseEvaluationOf(c, doc, caseNowOf(c)), sheetName: sheet.name })
        const stamp = { owner: 'rehearsal', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
        const source = residentSource(sheet, stamp, memory, 'engine')
        const disagree: string[] = []
        const painted = new Set<string>()
        try {
          for (const ref of refs) {
            const covering = c.markers.filter(m => m.sqref.split(' ').some(t => { const [a, z] = t.split(':'); return colOf(a!) === colOf(ref) && rowOf(ref) >= rowOf(a!) && rowOf(ref) <= rowOf(z ?? a!) }))
            const kind = covering.some(m => m.ruleType === 'iconSet') ? 'icon' : covering.some(m => m.ruleType === 'dataBar') ? 'bar' : undefined
            const f = obs.fill(ref)!
            const v = kind ? obs.visual!(ref, kind) : undefined
            const excel = { fill: f.fill, strike: f.strike, ...(v ? { visual: v.kind === 'none' ? 'none' : v.kind === 'bar' ? 'bar' : `icon:${(v as { cls: string }).cls}` } : {}) }
            const { row, col } = (() => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } })()
            const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
            const production = stage5TokenOf(lease.value.at(row, col))
            lease.release()
            const verdict = stage5CellAgreement(production, excel, truths.get(ref))
            if (!verdict.agrees) disagree.push(`${ref}: ${verdict.disagreement}`)
            if (production.unresolved.length === 0 && (production.fill !== 'none' || production.visual !== 'none')) painted.add(ref)
          }
        } finally { service.dispose() }
        expect(disagree).toEqual([])
        expect([...truths.values()].flatMap(t => t.unjustified)).toEqual([])
        // Non-vacuous: production paints the configurations the proposed readings unlock (scales over several ranges stay withheld).
        for (const ref of ['B11', 'B12', 'D13', 'D21', 'F3', 'F4', 'F12', 'F21']) expect(painted.has(ref), ref).toBe(true)
        for (const ref of ['B3', 'H3']) expect(painted.has(ref), ref).toBe(false)
      })
    } finally { injectCfFormulaReadings(null); injectCfFormulaSelections(null) }
  }, 300_000)
})

describe('rehearsing the proposed anchor-a1 registration on every measured Stage 5 case', () => {
  it('no measured cell would disagree with Excel or be withheld without a recorded justification (the older scope cases included)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const held = new Set([...VERIFIED_CAPABILITIES, MULTI, SINGLE])
    injectCfFormulaReadings({ [SINGLE]: 'anchor-a1', [MULTI]: ['rule-wide', 'per-range', 'anchor-a1'] })
    injectCfFormulaSelections({ ...VERIFIED_STAGE5_FORMULA_SELECTIONS, [MULTI]: 'anchor-a1' })
    try {
      const rows = await withInjectedCfEvidence(held, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => held.has(id) }))
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`)).toEqual([])
      // The older scope cases are compared under it: their multi-range scales are justified by scale:scope-disjoint.
      for (const id of ['stage5-cfvo-scope', 'stage5-cfvo-scope-v2', 'stage5-cfvo-formulas']) expect(rows.some(r => r.caseId === id && r.compared), id).toBe(true)
      expect(rows.some(r => r.caseId === 'stage5-cfvo-scope-v2' && r.justification.some(j => /scale:scope-disjoint not held/.test(j.reason)))).toBe(true)
      // stage5-cfvo-formulas K (=A2 on K2:K9) read from A1 wraps to a blank cell: a non-numeric bound, recorded.
      expect(rows.filter(r => r.caseId === 'stage5-cfvo-formulas' && /^K/.test(r.ref) && r.justification.some(j => /anchor-a1: =A2/.test(j.reason))).map(r => r.ref)).toEqual(['K2', 'K3', 'K4', 'K5', 'K6', 'K7', 'K8', 'K9'])
      // The superseded single-range record of stage5-cfvo-scope-v2 no longer names a refusal: production paints it.
      expect(rows.filter(r => r.caseId === 'stage5-cfvo-scope-v2' && /^I[2-7]$/.test(r.ref)).every(r => r.production.unresolved.length === 0 && r.production.fill !== 'none')).toBe(true)
    } finally { injectCfFormulaReadings(null); injectCfFormulaSelections(null) }
  }, 600_000)
})

describe('the recorded anchor-a1 refusals and supersessions are checked against the workbooks', () => {
  it('STAGE5_A1_BLANK_BOUND_REFUSALS: the wrapped window is blank in the saved workbook and the formula is the rule\'s bound', async () => {
    const { STAGE5_A1_BLANK_BOUND_REFUSALS, STAGE5_SUPERSEDED_RECORDS } = await import('../scripts/cf-refs/stage5-withhold')
    for (const e of STAGE5_A1_BLANK_BOUND_REFUSALS) {
      const c = manifest.cases.find(x => x.id === e.caseId)!
      const xml = new TextDecoder().decode(partsOf(c).get('xl/worksheets/sheet1.xml')!)
      expect(xml.includes(`r="${e.wrapped}"`), e.wrapped).toBe(false)
      expect(c.markers.some(m => m.sqref === e.sqref && (m.scale?.cfvo ?? []).some(b => (b.f ?? b.val) === e.formula)), e.sqref).toBe(true)
      // =A2 written on K2 (row 1, col 10, 0-based) read from A1 shifts by (-1, -10): row 0, column -10, wrapped to 16374.
      expect(e.wrapped).toBe('XEU1')
    }
    for (const e of STAGE5_SUPERSEDED_RECORDS) {
      const cfg = manifest.cases.find(x => x.id === e.caseId)!.formulaConfigurations!.find(x => x.id === e.configuration)!
      expect([cfg.record !== undefined, cfg.capabilityId]).toEqual([true, e.capabilityId])
    }
  })
})
