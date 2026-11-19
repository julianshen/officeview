// @vitest-environment node
/**
 * Follow-up to Task 39's re-deferral (orchestrator ruling 2026-11-19): stage5-bar-bounds-formula, a new case version for
 * bar:bounds-formula.
 *
 * stage5-cfvo-formulas I holds 5..100 against formula bounds $F$1/$F$2 = 20/80: a whole-scope cf-value-beyond-bounds
 * (Ruling S3-E), and its twin W7 is unmeasurable. This case keeps every value strictly inside 20..80, and every bar edge
 * either clear of the right-aligned value text or full width, under each reading (calibrated below on the committed
 * stage5-cfvo-formulas capture, the same 84 px cells and 20 pt rows). Its readings are compared with literal twins in the
 * same workbook, as every Stage 5 bar configuration is: num 20/80 (evaluated), num 0/0 (zero), min/max (ignored).
 *
 * The case has no capture yet, so the configuration is run through the real Stage 5 evaluator on a synthetic observation
 * (bars planted under one reading by an independent model of Excel's 84 px bar). This checks the family, not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_CASES, CF_STAGE5_GROUPS, CF_STAGE5_OWNERS, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  computeWorksheetRangeGeometry,
  decodePngSync,
  describeS5Visual,
  evaluateFormulaCapabilities,
  formulaControlsOf,
  formulaObservationOf,
  readZipPartsSync,
  resolveFormulaConfigurations,
  visualKindOf,
  type FormulaObservation,
  type S5Visual,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const ID = 'stage5-bar-bounds-formula'
const CAP = 'bar:bounds-formula'
const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const c = manifest.cases.find(x => x.id === ID)!
const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
const sheetXml = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml')!)
const valueOf = (ref: string): number => Number(new RegExp(`<c r="${ref}"[^>]*><v>([^<]+)</v>`).exec(sheetXml)![1])
const ROWS = [2, 3, 4, 5, 6, 7, 8, 9]
const VALUES = ROWS.map(r => valueOf(`B${r}`))

/**
 * Excel's solid 0/100 bar in a Stage 5 cell (84 px, left inset 2 px, an 81 px span), from the committed
 * stage5-cfvo-formulas capture: a value v under bounds lo/hi draws round(f * 81) px from x = 2, f = (v - lo) / (hi - lo);
 * nothing at or below lo; the full span at or above hi.
 */
const SPAN = 81
const barOf = (v: number, lo: number, hi: number): S5Visual => {
  const f = (v - lo) / (hi - lo)
  if (f <= 0) return { kind: 'none' }
  return { kind: 'bar', left: 2, right: 1 + Math.round(Math.min(1, f) * SPAN) }
}
/** The widest bar edge measured clear of the value text there (Y7 at 61 px; 73 and 75 px are under the text). */
const CLEAR_OF_TEXT_PX = 61
const FULL_PX = 82

describe('stage5-bar-bounds-formula: the case', () => {
  it('owns bar:bounds-formula (appended last), while stage5-cfvo-formulas keeps its configuration and capture', () => {
    expect(c.status).toBe('recipe-pending')
    expect(CF_STAGE5_OWNERS[CAP]).toBe(ID)
    expect(CF_STAGE5_CAPABILITIES.find(e => e.capabilityId === CAP)!.caseId).toBe(ID)
    expect(CF_STAGE5_CASES.at(-1)!.id).toBe(ID)
    expect(CF_STAGE5_CASES.find(x => x.id === 'stage5-cfvo-formulas')!.formulaConfigurations!.some(x => x.id === 'cfvo-bar-formula' && x.capabilityId === CAP)).toBe(true)
    expect(Object.values(CF_STAGE5_GROUPS).flat()).toContain(ID)
  })

  it('a formula-bound bar ($F$1/$F$2 = 20/80) and its literal twins on the same data: num 20/80, num 0/0, min/max', () => {
    expect([valueOf('F1'), valueOf('F2')]).toEqual([20, 80])
    for (const col of ['C', 'D', 'E']) expect(ROWS.map(r => valueOf(`${col}${r}`)), col).toEqual(VALUES)
    const bars = c.markers.filter(m => m.ruleType === 'dataBar' && m.extension === 'legacy').map(m => [m.capabilityId, m.sqref, m.bar!.cfvo])
    expect(bars).toEqual([
      [CAP, 'B2:B9', [{ type: 'formula', val: '$F$1' }, { type: 'formula', val: '$F$2' }]],
      ['formula:cfvo-bar-twin', 'C2:C9', [{ type: 'num', val: '20' }, { type: 'num', val: '80' }]],
      ['formula:cfvo-bar-twin', 'D2:D9', [{ type: 'num', val: '0' }, { type: 'num', val: '0' }]],
      ['formula:cfvo-bar-twin', 'E2:E9', [{ type: 'min' }, { type: 'max' }]],
    ])
    const cfg = c.formulaConfigurations!.find(x => x.capabilityId === CAP)!
    expect(cfg).toMatchObject({ id: 'bar-bounds-formula', observable: 'crop' })
    expect(Object.keys(cfg.hypotheses)).toEqual(['evaluated', 'zero', 'ignored'])
    for (const r of ROWS) expect([cfg.hypotheses.evaluated![`B${r}`], cfg.hypotheses.zero![`B${r}`], cfg.hypotheses.ignored![`B${r}`]]).toEqual([`twin:C${r}`, `twin:D${r}`, `twin:E${r}`])
    // S5-WITNESS: every num 20/80 twin draws a bar (all values above 20); the min/max twin is empty at the minimum only.
    expect(cfg.twinExpect).toMatchObject(Object.fromEntries(ROWS.map(r => [`C${r}`, 'visual'])))
    expect(cfg.twinExpect?.E2).toBe('empty')
  })

  it('every value lies strictly inside the formula bounds (no S3-E beyond-bounds cell), with 20 pt rows like the measured Stage 5 bars', () => {
    expect(VALUES).toEqual([28, 31, 35, 41, 46, 49, 52, 62])
    for (const v of VALUES) expect(v > 20 && v < 80, String(v)).toBe(true)
    expect(/<sheetFormatPr defaultRowHeight="20"/.test(sheetXml)).toBe(true)
  })

  it('every bar edge, under either reading, is clear of the value text or full width', () => {
    const lo = Math.min(...VALUES), hi = Math.max(...VALUES)
    for (const [a, b] of [[20, 80], [lo, hi]] as const) {
      for (const v of VALUES) {
        const bar = barOf(v, a, b)
        if (bar.kind === 'bar') expect(bar.right <= CLEAR_OF_TEXT_PX || bar.right === FULL_PX, `${v} under ${a}/${b}: right edge ${bar.right}`).toBe(true)
        // No length on a half pixel, where the model cannot say which way Excel rounds (see W4 below).
        const x = ((v - a) / (b - a)) * SPAN
        if (x > 0 && x < SPAN) expect(Math.abs((x % 1) - 0.5) >= 0.15, `${v} under ${a}/${b}: ${x} px`).toBe(true)
      }
    }
    // Largest edges: 58 px (62 under 20/80, and 52 under 28/62); 62 under min/max is the full span.
    expect(barOf(62, 20, 80)).toEqual({ kind: 'bar', left: 2, right: 58 })
    expect(barOf(52, 28, 62)).toEqual({ kind: 'bar', left: 2, right: 58 })
    expect(barOf(62, 28, 62)).toEqual({ kind: 'bar', left: 2, right: FULL_PX })
  })

  it('the bar model and the text margin match the committed stage5-cfvo-formulas capture', () => {
    const old = manifest.cases.find(x => x.id === 'stage5-cfvo-formulas')!
    const side = JSON.parse(readFileSync(resolve(REPO_ROOT, old.captures[0]!.sidecarPath), 'utf8')) as Record<string, unknown> & { colWidthsPt: number[]; rowHeightsPt: number[] }
    expect(side.colWidthsPt.slice(0, 26).every(w => w === 84)).toBe(true)
    expect(side.rowHeightsPt.every(h => h === 20)).toBe(true)
    const png = decodePngSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, old.captures[0]!.pngPath)))) ?? undefined
    const oldParts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, old.saved))))!
    const geom = computeWorksheetRangeGeometry(oldParts, old.range, { colWidthsPt: side.colWidthsPt, rowHeightsPt: side.rowHeightsPt })!
    const obs = formulaObservationOf(old, side, png, geom)
    const seen = (ref: string): string => { const v = obs.visual!(ref, 'bar'); return v ? describeS5Visual(v) : 'unmeasurable' }
    // Values 5,15,30,45,60,75,90,100 in rows 2..9. W: num 20/80; Y: min/max (5..100); X: num 0/0.
    const D = [5, 15, 30, 45, 60, 75, 90, 100]
    for (const [i, v] of D.entries()) {
      const r = i + 2
      // W4 (30 under 20/80) is a 13.5 px tie, which Excel drew as 13 px; the case keeps clear of ties.
      if (v !== 75 && v !== 30) expect(seen(`W${r}`), `W${r}`).toBe(describeS5Visual(barOf(v, 20, 80)))
      if (v !== 90) expect(seen(`Y${r}`), `Y${r}`).toBe(describeS5Visual(barOf(v, 5, 100)))
      expect(seen(`X${r}`), `X${r}`).toBe('bar:43-82')
    }
    expect(seen('W4')).toBe('bar:2-14')
    expect(seen('Y7')).toBe(`bar:2-${CLEAR_OF_TEXT_PX}`)
    // Edges at 75 px (W7, 75 under 20/80) and 73 px (Y8, 90 under 5/100) fall under the value text: unmeasurable.
    expect([barOf(75, 20, 80), barOf(90, 5, 100)]).toEqual([{ kind: 'bar', left: 2, right: 75 }, { kind: 'bar', left: 2, right: 73 }])
    expect([seen('W7'), seen('Y8')]).toEqual(['unmeasurable', 'unmeasurable'])
  })
})

describe('stage5-bar-bounds-formula: the configuration on a synthetic observation', () => {
  const lo = Math.min(...VALUES), hi = Math.max(...VALUES)
  const twinModel: Record<string, (v: number) => S5Visual> = {
    C: v => barOf(v, 20, 80),
    // Excel's num 0/0 twin draws the same right half whatever the value (stage5-cfvo-formulas X).
    D: () => ({ kind: 'bar', left: 43, right: 82 }),
    E: v => barOf(v, lo, hi),
  }
  const planted = { evaluated: 'C', zero: 'D', ignored: 'E' } as const
  const observe = (reading: keyof typeof planted): FormulaObservation => ({
    fill: () => ({ fill: 'none', strike: false }),
    visual(ref) {
      const col = /^[A-Z]+/.exec(ref)![0], row = Number(/\d+/.exec(ref)![0])
      if (!ROWS.includes(row) || !['B', 'C', 'D', 'E'].includes(col)) return { kind: 'none' }
      const v = valueOf(`${col}${row}`)
      return twinModel[col === 'B' ? planted[reading] : col]!(v)
    },
  })
  const judge = (reading: keyof typeof planted) =>
    evaluateFormulaCapabilities(resolveFormulaConfigurations(c, parts), formulaControlsOf(c), observe(reading), { kindOf: visualKindOf(c) })[CAP]!

  for (const reading of ['evaluated', 'ignored'] as const) {
    it(`planted ${reading}: bar:bounds-formula comes back as ${reading}, the one match`, () => {
      const ev = judge(reading)
      expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', reading])
    })
  }

  it('planted zero matches zero alone but is refused: the num 0/0 twin is degenerate, so it is no validated witness', () => {
    const ev = judge('zero')
    expect(ev.survivors).toEqual(['zero'])
    expect([ev.status, ev.code]).toEqual(['refused', 'formula-twin-blank'])
  })

  it('evaluated and ignored differ by at least 3 px, or bar against no bar, at every cell', () => {
    const apart = VALUES.filter(v => {
      const a = barOf(v, 20, 80), b = barOf(v, lo, hi)
      return a.kind !== b.kind || (a.kind === 'bar' && b.kind === 'bar' && Math.abs(a.right - b.right) >= 3)
    })
    expect(apart).toEqual(VALUES)
  })
})
