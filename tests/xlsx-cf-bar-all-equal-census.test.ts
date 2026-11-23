// @vitest-environment node
/**
 * The bar all-equal follow-up (brief cf-bar-allequal, Phase 1 task 1): pins the census of every captured data bar over an
 * all-equal (or single-numeric-cell) scope (scripts/cf-refs/bar-all-equal-census.ts) and the arithmetic of each closed
 * rule against Excel's px. The census tool itself registers nothing and is not reachable from production.
 *
 * Pinned as of its own run (fa933300, then 2a): the captures before stage3-bar-all-equal, under the published evidence
 * without the four ids that case registered (PRE_CAPTURE). The last block pins it after that capture (13bba185) and the
 * registration (docs/validation/xlsx-cf.md §14.z.20).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { barAllEqualCensus, type AllEqualCensus } from '../scripts/cf-refs/bar-all-equal-census'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'

const NEW_CASE = 'stage3-bar-all-equal'
const REGISTERED = ['bar:all-equal-distinct-bounds', 'bar:degenerate-equal-ltr', 'bar:degenerate-equal-rtl', 'bar:degenerate-equal-10-90']
const PRE_CAPTURE = [...VERIFIED_CAPABILITIES].filter(id => !REGISTERED.includes(id))
const fullManifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
/** The manifest as of the census's own run: every capture before stage3-bar-all-equal. */
const manifest = { ...fullManifest, cases: fullManifest.cases.filter(x => x.id !== NEW_CASE) } as CfManifest
const halfUp = (x: number): number => Math.floor(x + 0.5 + 1e-9)

/** Predicted bar length (px from the leading inset) on a span, for a closed rule at one census cell; 0 is no bar. */
function predict(rule: 'half' | 'full' | 'none' | 'evaluate', span: number, value: number, lo: number | undefined, hi: number | undefined): number | undefined {
  if (rule === 'none') return 0
  if (rule === 'full') return span
  if (rule === 'half') return halfUp(0.5 * span)
  // evaluate: (v - lo) / (hi - lo) on 0/100 lengths, only when lo < hi (no answer at lo == hi).
  if (lo === undefined || hi === undefined || !(lo < hi)) return undefined
  return halfUp(((value - lo) / (hi - lo)) * span)
}

let census: AllEqualCensus
const get = async (): Promise<AllEqualCensus> => (census ??= await withExactCfEvidence(PRE_CAPTURE, () => barAllEqualCensus(manifest)))

describe('bar all-equal census (every measured capture)', () => {
  it('scans every measured case with a bar; the cf-degenerate-scope refusals are exactly the all-equal rules (but the verified B, since 2a) plus num 0/0', async () => {
    const c = await get()
    const withBars = manifest.cases.filter(x => x.status === 'measured' && x.markers.some(m => m.ruleType === 'dataBar')).map(x => x.id).sort()
    expect(Object.keys(c.scanned).sort()).toEqual(withBars)
    expect(Object.values(c.scanned).reduce((a, b) => a + b, 0)).toBe(133)
    expect(c.degenerate.map(d => `${d.caseId} ${d.sqref} ${d.shape} ${d.position}`)).toEqual([
      'stage3-bar-degenerate C2 single-cell statistical (lo = hi = v)',
      'stage3-bar-geometry H2:H6 all-equal at hi',
      'stage3-bar-geometry I2:I6 all-equal at hi',
      'stage5-cfvo-formulas X2:X9 distinct lo == hi (0)',
      'stage5-bar-bounds-formula D2:D9 distinct lo == hi (0)',
    ])
  }, 300_000)

  it('records each all-equal rule: shape, Excel px and production (B paints the verified half since 2a; the rest cf-degenerate-scope)', async () => {
    const c = await get()
    expect(c.rules.map(r => [r.caseId, r.sqref, r.shape, r.pairing, r.bounds, r.lengths, r.gradient, r.border, r.direction, r.axis, r.value])).toEqual([
      ['stage3-bar-degenerate', 'B2:B6', 'all-equal', 'paired', 'min / max', '0/100', false, false, 'context', 'automatic', 42],
      ['stage3-bar-degenerate', 'C2', 'single-cell', 'paired', 'min / max', '0/100', false, false, 'context', 'automatic', 50],
      ['stage3-bar-geometry', 'H2:H6', 'all-equal', 'paired', 'num=0 / num=100', '0/100', false, false, 'context', 'automatic', 100],
      ['stage3-bar-geometry', 'I2:I6', 'all-equal', 'paired', 'num=0 / num=100', '0/100', false, false, 'context', 'automatic', 100],
    ])
    const px = (sqref: string) => c.rules.find(r => r.sqref === sqref)!.cells.map(x => `${x.ref} ${x.excel} ${x.excelLen}/${(x.cellWidthPx ?? 0) - 3}`)
    expect(px('B2:B6')).toEqual(['B2', 'B3', 'B4', 'B5', 'B6'].map(r => `${r} 86..174 89/177`))
    expect(px('C2')).toEqual(['C2 266..354 89/177'])
    expect(px('H2:H6')).toEqual(['H2', 'H3', 'H4', 'H5', 'H6'].map(r => `${r} 1459..1467 9/9`))
    expect(px('I2:I6')).toEqual(['I2', 'I3', 'I4', 'I5', 'I6'].map(r => `${r} 1471..1647 177/177`))
    // Before 2a every cell here was cf-degenerate-scope (fa933300); B2:B6 now paint 0..0.5 of the span (the 89 px Excel drew).
    for (const r of c.rules) for (const x of r.cells) expect(x.production, `${r.caseId}:${x.ref}`).toBe(r.sqref === 'B2:B6' ? 'painted 0..0.5' : 'cf-degenerate-scope')
  }, 300_000)

  it('closed rules against every captured cell (+-1 px): half fits only the statistical scopes; evaluate and full fit only num 0/100', async () => {
    const c = await get()
    const fits: Record<string, string[]> = {}
    for (const rule of ['half', 'full', 'none', 'evaluate'] as const) {
      fits[rule] = c.rules.filter(r => r.cells.every(x => {
        const [lo, hi] = r.bounds.startsWith('num') ? r.bounds.split(' / ').map(b => Number(b.split('=')[1])) : [undefined, undefined]
        const p = predict(rule, x.cellWidthPx! - 3, x.value, lo, hi)
        return p !== undefined && Math.abs(p - (x.excelLen ?? 0)) <= 1
      })).map(r => r.sqref)
    }
    expect(fits).toEqual({ half: ['B2:B6', 'C2'], full: ['H2:H6', 'I2:I6'], none: [], evaluate: ['H2:H6', 'I2:I6'] })
    // The misses, in px: half on H (5 vs 9) and I (89 vs 177); full on B (177 vs 89).
    expect([predict('half', 9, 100, 0, 100), predict('half', 177, 100, 0, 100), predict('full', 177, 42, undefined, undefined)]).toEqual([5, 89, 177])
  }, 300_000)
})

describe('bar all-equal census after the stage3-bar-all-equal capture (13bba185) and the registration', () => {
  it('only the shapes left withheld still refuse; earlier all-equal cells now paint as Excel drew them; the new case paints every all-equal rule', async () => {
    const c = await barAllEqualCensus(fullManifest)
    expect(Object.values(c.scanned).reduce((a, b) => a + b, 0)).toBe(147)
    expect(c.scanned[NEW_CASE]).toBe(14)
    // stage3-bar-geometry H and I (Excel full at hi) now paint under the registered `evaluate`.
    expect(c.degenerate.map(d => `${d.caseId} ${d.sqref}`)).toEqual(['stage3-bar-degenerate C2', 'stage5-cfvo-formulas X2:X9', 'stage5-bar-bounds-formula D2:D9'])
    const prod = (caseId: string, sqref: string) => [...new Set(c.rules.find(r => r.caseId === caseId && r.sqref === sqref)!.cells.map(x => x.production))]
    expect(prod('stage3-bar-geometry', 'H2:H6')).toEqual(['painted 0..1'])
    expect(prod('stage3-bar-geometry', 'I2:I6')).toEqual(['painted 0..1'])
    const mine = c.rules.filter(r => r.caseId === NEW_CASE).map(r => [r.sqref, prod(NEW_CASE, r.sqref).join(','), [...new Set(r.cells.map(x => x.excel))].join(',')])
    expect(mine).toEqual([
      ['B2:B4', 'painted 0..1', '86..277'],
      ['C2:C4', 'painted 0..0', 'none'],
      ['D2:D4', 'painted 0..0.3', '476..533'],
      ['E2:E4', 'painted 0..0.7', '671..804'],
      ['F2:F4', 'painted 0..0.05', '866..875'],
      ['G2:G4', 'painted 0..0.29', '1061..1116'],
      ['H2:H4', 'painted 0..0.8500000000000001', '1256..1418'],
      ['I2:I4', 'painted 0.75..1', '1595..1642'],
      ['L2:L4', 'painted 0..0.5', '2036..2131'],
      ['M2:M4', 'painted 0..0.5', '2231..2326'],
      ['N2:N4', 'painted 0..0.5', '2522..2617'],
      ['O2:O4', 'painted 0..0.5', '2621..2716'],
    ])
  }, 300_000)
})
