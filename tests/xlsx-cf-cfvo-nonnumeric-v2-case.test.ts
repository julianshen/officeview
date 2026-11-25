// @vitest-environment node
/**
 * The cfvo-nonnumeric confirmation case (brief 2026-11-21, Phase 1 Tasks 2 and 3; pre-capture ruling: one id per kind):
 * stage5-cfvo-nonnumeric-v2.
 *
 * stage5-cfvo-formulas L/M/N matched none of the uniform readings of formula:cfvo-nonnumeric (Excel painted nothing for
 * a text and an error bound, and read a blank one as 0). Production's per-kind treatment (0c572321, per-kind ids 67c3edb0)
 * post-dates that capture, so this case, built and captured after it, measures each kind on its own id
 * (formula:cfvo-nonnumeric-blank, -text, -empty-string, -error), each pre-registering withheld, zero and skipped: a surprise
 * in one kind refuses only that kind. Every reading of every id was planted before the capture on a synthetic observation
 * by an independent model (the Stage 2 sRGB truncate scale, 2 and 3 colours) and run through the real Stage 5 evaluator,
 * and those checks stay. The capture (16822e70) matched blank zero and text, the empty string and error withheld, each the
 * sole survivor; the four ids are registered and published, and the real registration is checked against the capture on
 * this case and every measured Stage 5 case.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  CF_STAGE5_CAPABILITIES, CF_STAGE5_CASES, CF_STAGE5_GROUPS, CF_STAGE5_NONDISCRIMINATING, CF_STAGE5_OWNERS, CF_STAGE5_UNSUPPORTED,
  NONNUM_V2_POSITIONS, NONNUM_V2_SOURCES, type CfManifest, type CfOracleCase,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { evaluateFormulaCapabilities, formulaControlsOf, readZipPartsSync, resolveFormulaConfigurations, type FormulaObservation } from '../scripts/cf-refs/verify'
import { caseEvaluationOf, caseNowOf, caseParseOptionsOf } from '../scripts/cf-refs/case-clock'
import { interpolateScaleColor } from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, injectCfFormulaSelections, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_CFVO_NONNUMERIC_KIND_IDS, CF_FORMULA_PRODUCTION_READINGS, injectCfFormulaReadings } from '../src/xlsx/cf/formulaReadings'
import { VERIFIED_STAGE5_FORMULA_SELECTIONS } from '../src/xlsx/cf/verifiedCapabilities'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

const REPO_ROOT = resolve(__dirname, '..')
const ID = 'stage5-cfvo-nonnumeric-v2'
const UMBRELLA = 'formula:cfvo-nonnumeric'
const TWIN = 'formula:cfvo-nonnumeric-v2-twin'
const TREATMENTS = ['withheld', 'zero', 'skipped'] as const
type Treatment = (typeof TREATMENTS)[number]
type Kind = 'blank' | 'text' | 'empty-string' | 'error'
const KINDS: readonly Kind[] = ['blank', 'text', 'empty-string', 'error']
type Position = (typeof NONNUM_V2_POSITIONS)[number]
type World = Readonly<Record<Kind, Treatment>>
/** The registration the old capture suggested (formed after it; never registered from it), which the capture 16822e70 matched. */
const SUGGESTED: World = { blank: 'zero', text: 'withheld', 'empty-string': 'withheld', error: 'withheld' }
const CAPTURED: World = SUGGESTED
const WORLDS: World[] = TREATMENTS.flatMap(b => TREATMENTS.flatMap(t => TREATMENTS.flatMap(s => TREATMENTS.map(e => ({ blank: b, text: t, 'empty-string': s, error: e })))))
const worldName = (w: World): string => KINDS.map(k => `${k}=${w[k]}`).join(' ')

const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (id = ID): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === id)
  if (!c) throw new Error(`${id} is not in the manifest (generate it with generate.ts --stage 5 --case ${id})`)
  return c
}
const partsOf = (c: CfOracleCase, file = c.saved): Map<string, Uint8Array> => readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, file))))!
const sheetXml = (c: CfOracleCase): string => new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
const valuesOf = (c: CfOracleCase): Map<string, number> => {
  const out = new Map<string, number>()
  for (const m of sheetXml(c).matchAll(/<c r="([A-Z]+\d+)"[^>]*><v>([^<]+)<\/v><\/c>/g)) if (Number.isFinite(Number(m[2]))) out.set(m[1]!, Number(m[2]))
  return out
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])

// ---- An independent model of each treatment's rendering (the Stage 2 sRGB truncate scale; Excel's default colours) ----
const LO = 0xfff8696b, MID = 0xffffeb84, HI = 0xff63be7b
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const scale2 = (v: number, lo: number, hi: number): string => (v <= lo ? hex(LO) : v >= hi ? hex(HI) : hex(interpolateScaleColor(LO, HI, (v - lo) / (hi - lo))))
const scale3 = (v: number, lo: number, mid: number, hi: number): string =>
  v <= lo ? hex(LO) : v >= hi ? hex(HI) : v <= mid ? hex(interpolateScaleColor(LO, MID, (v - lo) / (mid - lo))) : hex(interpolateScaleColor(MID, HI, (v - mid) / (hi - mid)))
/** What a formula-bound rule shows at a cell under one treatment, over column values min..max (undefined: no prediction). */
function rendering(t: Treatment, position: Position, v: number, min: number, max: number): string | undefined {
  if (t === 'withheld') return 'none'
  if (t === 'zero') return position === 'first' ? scale2(v, 0, max) : position === 'last' ? scale2(v, min, 0) : scale3(v, min, 0, max)
  return position === 'middle' ? undefined : scale2(v, min, max)
}
/** The configurations: source, position, the column the case put it in (C.. in source-major order) and its kind's id. */
const CONFIGS = NONNUM_V2_SOURCES.flatMap((src, i) => NONNUM_V2_POSITIONS.map((position, j) => ({ id: `nonnum-${src.id}-${position}`, src, kind: src.kind as Kind, cap: CF_CFVO_NONNUMERIC_KIND_IDS[src.kind as Kind], position, col: String.fromCharCode(67 + i * 3 + j) })))
const ROWS = [2, 3, 4, 5, 6]

describe('stage5-cfvo-nonnumeric-v2: the case', () => {
  it('holds one evidence id per kind (appended last), each published with its matched reading; the umbrella stays a recorded refusal', () => {
    const c = caseOf()
    expect(c.status).toBe('measured')
    for (const k of KINDS) {
      const id = CF_CFVO_NONNUMERIC_KIND_IDS[k]
      expect(NONNUM_V2_SOURCES.find(s => s.kind === k)!.cap).toBe(id)
      expect(CF_STAGE5_CAPABILITIES.find(e => e.capabilityId === id)!.caseId, id).toBe(ID)
      expect(CF_STAGE5_UNSUPPORTED[id], id).toBeUndefined()
      // Registered as the selection list after the capture (16822e70) matched one sole survivor per id.
      expect(CF_FORMULA_PRODUCTION_READINGS[id], id).toEqual(['withheld', 'zero', 'skipped'])
      expect([VERIFIED_CAPABILITIES.has(id), VERIFIED_STAGE5_FORMULA_SELECTIONS[id]], id).toEqual([true, CAPTURED[k]])
    }
    // The umbrella is superseded by the per-kind ids: no longer registered, never published.
    expect(VERIFIED_CAPABILITIES.has(UMBRELLA)).toBe(false)
    expect(CF_STAGE5_OWNERS[UMBRELLA]).toBeUndefined()
    expect(CF_STAGE5_CAPABILITIES.find(e => e.capabilityId === UMBRELLA)!.caseId).toBe('stage5-cfvo-formulas')
    expect(CF_STAGE5_UNSUPPORTED[UMBRELLA]).toMatch(/Superseded per kind/)
    expect(CF_FORMULA_PRODUCTION_READINGS[UMBRELLA]).toBeUndefined()
    // Appended last at its time; the percent-order follow-up's icon case comes after it.
    expect(CF_STAGE5_CASES.slice(-2).map(x => x.id)).toEqual([ID, 'stage5-icon-percent-order'])
    expect(CF_STAGE5_GROUPS['stage5-cfvo-formulas']).toContain(ID)
    expect(CF_STAGE5_NONDISCRIMINATING[TWIN]).toBeDefined()
    expect(caseOf('stage5-cfvo-formulas').status).toBe('measured')
  })

  it('the rules: one colour scale per source and position on its kind\'s id, then four literal twins; nothing else covers a designated cell', () => {
    const c = caseOf()
    const own = c.markers.filter(m => m.capabilityId !== TWIN && m.capabilityId !== 'formula:positive-control')
    const bound = (x: (typeof CONFIGS)[number]): string => `formula:$${colOf(x.src.ref)}$${rowOf(x.src.ref)}`
    expect(own.map(m => [m.capabilityId, m.sqref, m.scale!.cfvo.map(x => x.type === 'formula' ? `formula:${x.f ?? x.val}` : x.type), m.scale!.colors.length])).toEqual(
      CONFIGS.map(x => [x.cap, `${x.col}2:${x.col}6`, x.position === 'first' ? [bound(x), 'max'] : x.position === 'last' ? ['min', bound(x)] : ['min', bound(x), 'max'], x.position === 'middle' ? 3 : 2]))
    expect(c.markers.filter(m => m.capabilityId === TWIN).map(m => [m.sqref, m.scale!.cfvo.map(x => x.val ?? x.type).join(',')])).toEqual([
      ['P2:P6', '0,max'], ['Q2:Q6', 'min,max'], ['R2:R6', 'min,0'], ['S2:S6', 'min,0,max'],
    ])
    for (const e of c.expected.cells) expect(c.markers.filter(m => m.extension === 'legacy' && colOf(m.sqref) === colOf(e.ref) && rowOf(e.ref) >= 2 && rowOf(e.ref) <= 6).length, e.ref).toBeLessThanOrEqual(1)
    expect(c.expected.cells.length).toBe(160)
  })

  it('the bound sources: A1 absent (blank), A2 "abc", A3 ="" with its cached empty string, A4 #N/A; nothing else in column A', () => {
    const xml = sheetXml(caseOf())
    expect(xml).not.toMatch(/<c r="A1"/)
    expect(xml).toContain('<c r="A2" t="inlineStr"><is><t xml:space="preserve">abc</t></is></c>')
    expect(xml).toContain('<c r="A3" t="str"><f>&quot;&quot;</f><v></v></c>')
    expect(xml).toContain('<c r="A4" t="e"><v>#N/A</v></c>')
    expect([...xml.matchAll(/<c r="A(\d+)"/g)].map(m => m[1])).toEqual(['2', '3', '4'])
  })

  it('each configuration names withheld, zero and skipped on its kind\'s id; a midpoint omits skipped, with its reason', () => {
    const cfgs = caseOf().formulaConfigurations!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.observable, x.refs.join(','), x.record ?? ''])).toEqual(CONFIGS.map(x => [x.id, x.cap, 'fill', ROWS.map(r => `${x.col}${r}`).join(','), '']))
    for (const x of CONFIGS) {
      const cfg = cfgs.find(y => y.id === x.id)!
      expect(Object.keys(cfg.hypotheses), x.id).toEqual(x.position === 'middle' ? ['withheld', 'zero'] : ['withheld', 'zero', 'skipped'])
      expect(Object.keys(cfg.omitted ?? {}), x.id).toEqual(x.position === 'middle' ? ['skipped'] : [])
      for (const [t, preds] of Object.entries(cfg.hypotheses)) {
        const twin = t === 'zero' ? { first: 'P', last: 'R', middle: 'S' }[x.position] : 'Q'
        for (const r of ROWS) expect(preds[`${x.col}${r}`], `${x.id} ${t}`).toBe(t === 'withheld' ? 'none' : `twin:${twin}${r}`)
      }
    }
  })

  it('every pair of treatments is separated at >= 3 decisive cells of every configuration, beyond a 0.1-unit channel error', () => {
    const values = valuesOf(caseOf())
    const channels = (argb: string): number[] => [2, 4, 6].map(k => parseInt(argb.slice(k, k + 2), 16))
    for (const x of CONFIGS) {
      const vs = ROWS.map(r => values.get(`${x.col}${r}`)!)
      const min = Math.min(...vs), max = Math.max(...vs)
      const predicted = TREATMENTS.filter(t => rendering(t, x.position, vs[0]!, min, max) !== undefined)
      expect(predicted.length, x.id).toBe(x.position === 'middle' ? 2 : 3)
      for (const w of predicted) {
        // A decisive cell for winner w: its rendering differs from every other predicted treatment's there.
        const decisive = ROWS.filter((_, i) => predicted.filter(t => t !== w).every(t => rendering(t, x.position, vs[i]!, min, max) !== rendering(w, x.position, vs[i]!, min, max)))
        expect(decisive.length, `${x.id} ${w}`).toBeGreaterThanOrEqual(3)
        for (const r of decisive) {
          const v = values.get(`${x.col}${r}`)!
          const a = rendering(w, x.position, v, min, max)!
          for (const t of predicted.filter(y => y !== w)) {
            const b = rendering(t, x.position, v, min, max)!
            if (a !== 'none' && b !== 'none') expect(Math.max(...channels(a).map((c, k) => Math.abs(c - channels(b)[k]!))), `${x.id} ${r}: ${a} / ${b}`).toBeGreaterThanOrEqual(2)
          }
        }
      }
      for (const v of vs) for (const [lo, hi] of [[0, max], [min, 0], [min, max]] as const) if (v > lo && v < hi) expect(Number.isInteger(((v - lo) / (hi - lo)) * 1048576), `${x.id} ${v}`).toBe(true)
    }
  })
})

/** A planted observation: each formula column renders its kind's treatment, twins their literal rule, every other cell none. */
function plant(values: Map<string, number>, w: World, override: (ref: string) => string | undefined = () => undefined): FormulaObservation {
  const twinRender: Record<string, (v: number) => string> = { P: v => scale2(v, 0, 32), Q: v => scale2(v, -32, 32), R: v => scale2(v, -32, 0), S: v => scale3(v, -32, 0, 32) }
  return {
    fill(ref) {
      const v = values.get(ref)
      if (v === undefined || rowOf(ref) < 2 || rowOf(ref) > 6) return { fill: 'none', strike: false }
      const o = override(ref)
      if (o !== undefined) return { fill: o, strike: false }
      const x = CONFIGS.find(y => y.col === colOf(ref))
      // A midpoint under skipped is unpredicted: the planted world shows nothing there.
      if (x) return { fill: rendering(w[x.kind], x.position, v, -32, 32) ?? 'none', strike: false }
      const tw = twinRender[colOf(ref)]
      return { fill: tw ? tw(v) : 'none', strike: false }
    },
  }
}
const judge = (obs: FormulaObservation) => {
  const c = caseOf()
  return evaluateFormulaCapabilities(resolveFormulaConfigurations(c, partsOf(c)), formulaControlsOf(c), obs)
}

describe('stage5-cfvo-nonnumeric-v2: each kind discriminates on its own, on a synthetic observation', () => {
  it('every reading of every id, in all 81 planted worlds, comes back as that id\'s one match', () => {
    const c = caseOf(), values = valuesOf(c)
    for (const w of WORLDS) {
      const ev = judge(plant(values, w))
      for (const k of KINDS) {
        const e = ev[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!
        expect([e.status, e.matched], `${worldName(w)}: ${k} (${e.reason})`).toEqual(['discriminating', w[k]])
      }
      expect(ev[UMBRELLA], 'no umbrella configuration in this case').toBeUndefined()
    }
  }, 300_000)
  it('a surprise in one kind refuses only that kind: ="" read like a blank in one position, or text split by position', () => {
    const c = caseOf(), values = valuesOf(c)
    const zeroFirst = (ref: string): string => scale2(values.get(ref)!, 0, 32)
    // ="" (I, first) shows the zero rendering while its other positions paint nothing.
    const emptySplit = judge(plant(values, SUGGESTED, ref => (colOf(ref) === 'I' ? zeroFirst(ref) : undefined)))
    expect([emptySplit[CF_CFVO_NONNUMERIC_KIND_IDS['empty-string']]!.status, (emptySplit[CF_CFVO_NONNUMERIC_KIND_IDS['empty-string']] as { code?: string }).code]).toEqual(['refused', 'formula-hypothesis-none'])
    for (const k of ['blank', 'text', 'error'] as const) expect([emptySplit[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!.status, emptySplit[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!.matched], k).toEqual(['discriminating', SUGGESTED[k]])
    // Text (F, first) shows the zero rendering while its other positions paint nothing.
    const textSplit = judge(plant(values, SUGGESTED, ref => (colOf(ref) === 'F' ? zeroFirst(ref) : undefined)))
    expect(textSplit[CF_CFVO_NONNUMERIC_KIND_IDS.text]!.status).toBe('refused')
    for (const k of ['blank', 'empty-string', 'error'] as const) expect([textSplit[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!.status, textSplit[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!.matched], k).toEqual(['discriminating', SUGGESTED[k]])
  })
})

/**
 * Production as it would run with each kind in `held` registered (the selection list) and selected to `w`'s treatment
 * (the test seams) and held; every other per-kind id is registered but not held (a kind whose evidence failed).
 */
async function withProposed<T>(w: World, held: readonly Kind[], fn: (published: (id: string) => boolean) => Promise<T>): Promise<T> {
  const ids = held.map(k => CF_CFVO_NONNUMERIC_KIND_IDS[k])
  const set = new Set([...[...VERIFIED_CAPABILITIES].filter(id => !Object.values(CF_CFVO_NONNUMERIC_KIND_IDS).includes(id)), ...ids])
  injectCfFormulaReadings(Object.fromEntries(KINDS.map(k => [CF_CFVO_NONNUMERIC_KIND_IDS[k], TREATMENTS])))
  injectCfFormulaSelections({ ...VERIFIED_STAGE5_FORMULA_SELECTIONS, ...Object.fromEntries(held.map(k => [CF_CFVO_NONNUMERIC_KIND_IDS[k], w[k]])) })
  try { return await withExactCfEvidence([...set], () => fn(id => set.has(id))) } finally { injectCfFormulaReadings(null); injectCfFormulaSelections(null) }
}

/** The production counterfactual over this case on the planted world `w`: disagreements, unjustified refusals, painted cells. */
async function rehearse(w: World, held: readonly Kind[] = KINDS) {
  const { stage5CounterfactualOf, stage5CellAgreement, stage5TokenOf } = await import('../scripts/cf-refs/stage5-withhold')
  const c = caseOf()
  const obs = plant(valuesOf(c), w)
  return withProposed(w, held, async published => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))), caseParseOptionsOf(c))
    const sheet = doc.sheets[0]!
    const refs = c.expected.cells.map(e => e.ref)
    const truths = await stage5CounterfactualOf({ c, doc, sheet, refs, now: caseNowOf(c), published })
    const memory = createCfMemory()
    const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation: caseEvaluationOf(c, doc, caseNowOf(c)), sheetName: sheet.name })
    const stamp = { owner: 'rehearsal', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
    const source = residentSource(sheet, stamp, memory, 'engine')
    const disagree: string[] = [], painted = new Set<string>(), withheld = new Set<string>()
    try {
      for (const ref of refs) {
        const f = obs.fill(ref)!
        const { row, col } = { row: rowOf(ref) - 1, col: [...colOf(ref)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
        const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
        const production = stage5TokenOf(lease.value.at(row, col))
        lease.release()
        const verdict = stage5CellAgreement(production, { fill: f.fill, strike: false }, truths.get(ref))
        if (!verdict.agrees) disagree.push(`${ref}: ${verdict.disagreement}`)
        if (production.unresolved.length > 0) withheld.add(ref)
        else if (production.fill !== 'none') painted.add(ref)
      }
    } finally { service.dispose() }
    return { disagree, unjustified: [...truths.entries()].flatMap(([r, x]) => x.unjustified.map(u => `${r}: ${u}`)), painted, withheld }
  })
}
const resolves = (t: Treatment | undefined, position: Position): boolean => t === 'zero' || (t === 'skipped' && position !== 'middle')
/**
 * The num-segment follow-up (owner ruling A): a resolved midpoint of 0 puts rows 4 and 5 (8 and 16) strictly inside the upper
 * segment 0..32 at whole step products, which production withholds (cf-unmeasured-upper-segment-whole-product, justified by
 * the shape recomputed from the sheet). The resolving checks read row 3 (-8, lower segment) for a midpoint instead.
 */
const upperWhole = (position: Position, r: number): boolean => position === 'middle' && (r === 4 || r === 5)
const probeRow = (position: Position): number => (position === 'middle' ? 3 : 4)

describe('stage5-cfvo-nonnumeric-v2: the per-kind registration against production on planted worlds (every reading of every id)', () => {
  it('the registered and captured world: no unjustified refusal, no disagreement; blank paints, text, ="" and error withhold', async () => {
    const { disagree, unjustified, painted, withheld } = await rehearse(SUGGESTED)
    expect(disagree).toEqual([])
    expect(unjustified).toEqual([])
    for (const x of CONFIGS) for (const r of ROWS) expect([painted.has(`${x.col}${r}`), withheld.has(`${x.col}${r}`)], `${x.id} ${r}`).toEqual(x.kind === 'blank' && !upperWhole(x.position, r) ? [true, false] : [false, true])
    for (const col of ['P', 'Q', 'R', 'S']) for (const r of ROWS) expect(painted.has(`${col}${r}`), `${col}${r}`).toBe(!(col === 'S' && (r === 4 || r === 5)))
  }, 300_000)

  it('every one of the 81 worlds, registered in full: no unjustified refusal, no disagreement; each resolving bound paints', async () => {
    for (const w of WORLDS) {
      const { disagree, unjustified, painted } = await rehearse(w)
      expect(disagree, worldName(w)).toEqual([])
      expect(unjustified, worldName(w)).toEqual([])
      for (const x of CONFIGS) expect(painted.has(`${x.col}${probeRow(x.position)}`), `${worldName(w)} ${x.id}`).toBe(resolves(w[x.kind], x.position))
    }
  }, 600_000)

  it('a kind whose evidence failed (registered, not published) stays withheld, justified; the others paint as registered', async () => {
    for (const failed of KINDS) {
      const w: World = { blank: 'zero', text: 'zero', 'empty-string': 'zero', error: 'zero' }
      const { disagree, unjustified, painted, withheld } = await rehearse(w, KINDS.filter(k => k !== failed))
      expect([disagree, unjustified], failed).toEqual([[], []])
      for (const x of CONFIGS) expect(x.kind === failed ? withheld.has(`${x.col}${probeRow(x.position)}`) : painted.has(`${x.col}${probeRow(x.position)}`), `${failed}: ${x.id}`).toBe(true)
    }
  }, 300_000)
})

describe('the registration against the capture: every measured Stage 5 case', () => {
  it('the capture matched one sole survivor per id: blank zero; text, the empty string and error withheld', () => {
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, caseOf().captures[0]!.sidecarPath!), 'utf8')) as { formulaEvaluation: Record<string, { status: string; matched?: string; survivors?: string[] }> }
    for (const k of KINDS) {
      const e = sidecar.formulaEvaluation[CF_CFVO_NONNUMERIC_KIND_IDS[k]]!
      expect([e.status, e.matched, e.survivors], k).toEqual(['discriminating', CAPTURED[k], [CAPTURED[k]]])
    }
  })

  it('no measured cell disagrees with Excel or is withheld without a recorded justification; this case and stage5-cfvo-formulas K, M paint as Excel, L, N withhold', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const rows = await stage5CaseRecords(manifest, REPO_ROOT)
    // This case, compared with its capture: blank bounds paint (agreeing), the others withhold with their justification.
    for (const x of CONFIGS) for (const r of ROWS) {
      const row = rows.find(y => y.caseId === ID && y.ref === `${x.col}${r}`)!
      expect([row.compared, row.agrees, row.production.unresolved.length > 0], `${x.id} ${r}`).toEqual([true, true, x.kind !== 'blank' || upperWhole(x.position, r)])
      if (x.kind === 'blank' && upperWhole(x.position, r)) expect(row.justification.map(j => j.reason).join('; '), `${x.id} ${r}`).toContain('stage2-scale-num-segment withhold')
      if (x.kind !== 'blank') expect(row.justification.map(j => j.reason).join('; '), `${x.id} ${r}`).toContain(`${CF_CFVO_NONNUMERIC_KIND_IDS[x.kind]} = withheld withholds ${x.kind}`)
    }
    expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`)).toEqual([])
    const old = (ref: string) => rows.find(r => r.caseId === 'stage5-cfvo-formulas' && r.ref === ref)!
    for (const col of ['K', 'M']) for (let r = 2; r <= 9; r++) expect([old(`${col}${r}`).compared, old(`${col}${r}`).agrees, old(`${col}${r}`).production.unresolved], `${col}${r}`).toEqual([true, true, []])
    for (const [col, kind] of [['L', 'text'], ['N', 'error']] as const) for (let r = 2; r <= 9; r++) {
      const row = old(`${col}${r}`)
      expect([row.compared, row.agrees, row.production.unresolved.length > 0], `${col}${r}`).toEqual([true, true, true])
      expect(row.justification.map(j => j.reason).join('; '), `${col}${r}`).toContain(`formula:cfvo-nonnumeric-${kind} = withheld withholds ${kind}`)
    }
    expect(rows.some(r => r.caseId === 'stage5-cfvo-formulas' && r.justification.some(j => /anchor-a1: =A2/.test(j.reason)))).toBe(false)
  }, 600_000)
})

describe('the recorded non-numeric bounds are checked against the workbooks', () => {
  it('STAGE5_NONNUMERIC_BOUNDS: each bound is its rule\'s formula point at that position, and its cell holds that kind in the saved workbook', async () => {
    const { STAGE5_NONNUMERIC_BOUNDS } = await import('../scripts/cf-refs/stage5-withhold')
    for (const e of STAGE5_NONNUMERIC_BOUNDS) {
      const c = caseOf(e.caseId)
      const m = c.markers.find(x => x.sqref === e.sqref && x.extension === 'legacy')!
      const cfvo = m.scale!.cfvo
      const i = e.position === 'first' ? 0 : e.position === 'last' ? cfvo.length - 1 : 1
      expect([cfvo.length === 3 || e.position !== 'middle', cfvo[i]!.type, cfvo[i]!.f ?? cfvo[i]!.val], `${e.caseId} ${e.sqref}`).toEqual([true, 'formula', e.formula])
      const cell = /^\$([A-Z]+)\$(\d+)$/.exec(e.formula)!
      const xml = new TextDecoder().decode(partsOf(c).get('xl/worksheets/sheet1.xml')!)
      const at = new RegExp(`<c r="${cell[1]}${cell[2]}"[^>]*?(?:/>|>.*?</c>)`).exec(xml)?.[0]
      if (e.kind === 'blank') expect(at, `${e.caseId} ${e.formula}`).toBeUndefined()
      else if (e.kind === 'error') expect(at, `${e.caseId} ${e.formula}`).toMatch(/ t="e"/)
      else if (e.kind === 'empty-string') expect(at, `${e.caseId} ${e.formula}`).toMatch(/ t="str"><f>[^<]*<\/f><v(?:\/>|><\/v>)/)
      else expect(at, `${e.caseId} ${e.formula}`).toMatch(/ t="(s|str|inlineStr)"[^>]*>(?!.*<v><\/v>)/)
    }
    // Complete: every umbrella or per-kind marker of a Stage 5 case is listed.
    const ids = new Set([UMBRELLA, ...Object.values(CF_CFVO_NONNUMERIC_KIND_IDS)])
    const listed = new Set(STAGE5_NONNUMERIC_BOUNDS.map(e => `${e.caseId}|${e.sqref}`))
    for (const c of manifest.cases.filter(x => x.stage === 5)) for (const m of c.markers.filter(x => ids.has(x.capabilityId) && x.extension === 'legacy')) expect(listed.has(`${c.id}|${m.sqref}`), `${c.id} ${m.sqref}`).toBe(true)
  })

  it('a withheld treatment, skipped at a midpoint or no held treatment justifies the refusal; a resolving one does not', async () => {
    const { stage5JustifiedRefusals, stage5WithholdJustification } = await import('../scripts/cf-refs/stage5-withhold')
    const reasons = (id: string, ref: string, w: World, held: readonly Kind[] = KINDS): Promise<string[]> => withProposed(w, held, async published => stage5JustifiedRefusals(caseOf(id), ref, published).map(r => r.reason))
    expect(await reasons('stage5-cfvo-formulas', 'L3', SUGGESTED)).toEqual(['$F$3 is text: formula:cfvo-nonnumeric-text = withheld withholds text: the rule paints nothing and keeps its slot taint'])
    expect(await reasons('stage5-cfvo-formulas', 'M3', SUGGESTED)).toEqual([])
    expect(await reasons('stage5-cfvo-formulas', 'K3', SUGGESTED)).toEqual([])
    const blankWithheld: World = { blank: 'withheld', text: 'zero', 'empty-string': 'zero', error: 'zero' }
    expect(await reasons('stage5-cfvo-formulas', 'M3', blankWithheld)).toEqual([expect.stringMatching(/formula:cfvo-nonnumeric-blank = withheld withholds blank/)])
    expect(await reasons('stage5-cfvo-formulas', 'K3', blankWithheld)).toEqual([expect.stringMatching(/^anchor-a1: =A2.*formula:cfvo-nonnumeric-blank = withheld withholds blank/)])
    expect(await reasons('stage5-cfvo-formulas', 'L3', blankWithheld)).toEqual([])
    const skipped: World = { blank: 'skipped', text: 'skipped', 'empty-string': 'skipped', error: 'skipped' }
    expect(await reasons(ID, 'E4', skipped)).toEqual([expect.stringMatching(/skips blank at a midpoint/)])
    expect(await reasons(ID, 'C4', skipped)).toEqual([])
    // A failed kind (registered, not published): its rules are justified by its unheld id.
    expect(await reasons(ID, 'I4', SUGGESTED, ['blank', 'text', 'error'])).toEqual(['formula:cfvo-nonnumeric-empty-string not held'])
    // The committed published set: text and the empty string are justified by their published withheld readings; nothing
    // registered at all would leave the marker's own id as the justification (`<id> not held`).
    expect(stage5WithholdJustification(caseOf('stage5-cfvo-formulas'), 'L3')).toEqual([`${UMBRELLA}/colorScale@p9: $F$3 is text: formula:cfvo-nonnumeric-text = withheld withholds text: the rule paints nothing and keeps its slot taint`])
    expect(stage5WithholdJustification(caseOf(), 'I4')).toEqual(['formula:cfvo-nonnumeric-empty-string/colorScale@p7: $A$3 is the empty string: formula:cfvo-nonnumeric-empty-string = withheld withholds empty-string: the rule paints nothing and keeps its slot taint'])
    expect(stage5WithholdJustification(caseOf(), 'C4')).toEqual([])
    expect(stage5WithholdJustification(caseOf(), 'I4', id => VERIFIED_CAPABILITIES.has(id) && id !== CF_CFVO_NONNUMERIC_KIND_IDS['empty-string'])).toEqual(['formula:cfvo-nonnumeric-empty-string/colorScale@p7: formula:cfvo-nonnumeric-empty-string not held'])
  }, 300_000)
})
