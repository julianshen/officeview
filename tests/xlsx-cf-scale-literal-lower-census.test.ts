// @vitest-environment node
/**
 * The literal-lower follow-up (brief cf-literal-lower), Phase 1 task 1: the census of every captured interior cell of a
 * 3-colour LOWER segment bounded by two literal num cfvo, Stage 2 and the Stage 5 painted colour-scale cells
 * (scripts/cf-refs/literal-lower-census.ts; docs/validation/xlsx-cf-recipes-stage2.md §12.1). Production's lower-segment
 * arithmetic as now registered (raw floor at guard-boundary products, `interpolateScaleColor` truncation elsewhere) equals
 * Excel's DisplayFormat fill at every one of them, the 64 the num-segment withhold covers (f·2^20 not an integer) included.
 * This checks the census, not Excel. The census as run before stage2-scale-literal-lower was captured is pinned first
 * (recipe §12.1); with that capture (a3f85309) included it is pinned at the end (recipe §12.7): still exact.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { countBy, literalLowerCensusOf, productionLowerColour, type LiteralLowerRow } from '../scripts/cf-refs/literal-lower-census'

const HEAVY = 300_000
let census: Awaited<ReturnType<typeof literalLowerCensusOf>>
let rows: LiteralLowerRow[] = []
let withheld: LiteralLowerRow[] = []
beforeAll(async () => {
  // The census as run before stage2-scale-literal-lower was captured (recipe §12.1); the capture is pinned separately below.
  census = await literalLowerCensusOf(undefined, undefined, { include: 'pre-capture' })
  rows = census.rows
  withheld = rows.filter(r => !r.f20)
}, HEAVY)
const table = (rs: readonly LiteralLowerRow[], key: (r: LiteralLowerRow) => string): string[] => countBy(rs, key).map(t => `${t.key} ${t.cells} ${t.exact}`)

describe('the literal-lower census (every captured interior cell of a literal-num 3-colour lower segment)', () => {
  it('1421 captured segment cells, no harness fault; 70 literal-num lower-segment interior cells, 64 of them in the num-segment shape (f·2^20 not an integer), 7 endpoints or clamped', () => {
    expect(census.collection.cells.length).toBe(1421)
    expect(census.collection.faults).toEqual([])
    expect(rows.length).toBe(70)
    expect(withheld.length).toBe(64)
    expect(census.endpoints).toBe(7)
    // The harness evaluates the 64 as production withholds them (the num-segment shape), the 6 others as painted.
    expect(withheld.every(r => r.status === 'withheld')).toBe(true)
    expect(rows.filter(r => r.f20).every(r => r.status === 'confirmed')).toBe(true)
  })

  it('production\'s lower-segment arithmetic equals Excel\'s fill at all 70 (all 64 withheld cells included): no disagreement, so the case is not blocked', () => {
    expect(rows.filter(r => r.lower !== r.excel)).toEqual([])
  })

  it('the 64 by case, span, colour order, palette kind and product class', () => {
    expect(table(withheld, r => r.caseId)).toEqual(['stage2-scale-num-segment 48 48', 'stage2-scale-raw-floor 4 4', 'stage2-scale-scope-disjoint 6 6', 'stage2-scale-upper-whole 6 6'])
    expect(table(withheld, r => `${r.span} (${r.lo}..${r.hi})`)).toEqual(['210 (0..210) 4 4', '210 (7..217) 52 52', '30 (0..30) 3 3', '30 (30..60) 2 2', '45 (0..45) 3 3'])
    expect(table(withheld, r => `${r.order} ${r.palette}`)).toEqual(['asc FFF8696B>FFFFEB84 8 8', 'both FF14F028>FFE61EFA 34 34', 'both FFE61EFA>FF14F028 22 22'])
    expect(table(withheld, r => r.kind)).toEqual(['distinct 8 8', 'p-line 56 56'])
    expect(table(withheld, r => r.cls)).toEqual(['below 4 4', 'clear 30 30', 'integer 30 30'])
    expect(table(withheld, r => `${r.kind} ${r.cls}`)).toEqual(['distinct clear 8 8', 'p-line below 4 4', 'p-line clear 22 22', 'p-line integer 30 30'])
    // Gaps the new case fills: no cell above an integer; distinct spans only ascending and clear; below-integer cells only at span 210.
    expect(withheld.some(r => r.cls === 'above')).toBe(false)
    expect(withheld.filter(r => r.cls === 'below').map(r => `${r.caseId}:${r.ref}`)).toEqual(['stage2-scale-raw-floor:F6', 'stage2-scale-raw-floor:F7', 'stage2-scale-raw-floor:F8', 'stage2-scale-raw-floor:F9'])
    expect(withheld.every(r => r.lo >= 0)).toBe(true)
  })

  it('productionLowerColour: raw floor at a guard-boundary product (57/210 on the P line), truncation elsewhere', () => {
    expect(productionLowerColour(0xff14f028, 0xffe61efa, 57, 0, 210, 57 / 210)).toBe('FF4CB860') // 57/210·210 = 56.99999999999999: step 56
    expect(productionLowerColour(0xff14f028, 0xffe61efa, 105, 0, 210, 0.5)).toBe('FF7D8791')
  })
})

describe('after the capture of stage2-scale-literal-lower (a3f85309): the census on every capture (recipe §12.7)', () => {
  let post: Awaited<ReturnType<typeof literalLowerCensusOf>>
  beforeAll(async () => { post = await literalLowerCensusOf() }, HEAVY)

  it('1589 segment cells, no harness fault; 130 literal-num lower-segment interior cells (60 from the capture), 114 in the num-segment shape (50 from it)', () => {
    expect(post.collection.cells.length).toBe(1589)
    expect(post.collection.faults).toEqual([])
    expect(post.rows.length).toBe(130)
    expect(post.rows.filter(r => r.caseId === 'stage2-scale-literal-lower').length).toBe(60)
    expect(post.rows.filter(r => !r.f20).length).toBe(114)
    expect(post.rows.filter(r => !r.f20 && r.caseId === 'stage2-scale-literal-lower').length).toBe(50)
  })

  it('production\'s lower-segment arithmetic still equals Excel\'s fill at all 130; the 114 of the shape by palette kind and class', () => {
    expect(post.rows.filter(r => r.lower !== r.excel)).toEqual([])
    const shape = post.rows.filter(r => !r.f20)
    expect(table(shape, r => `${r.kind} ${r.cls}`)).toEqual(['distinct above 4 4', 'distinct below 3 3', 'distinct clear 11 11', 'distinct integer 6 6', 'p-line above 12 12', 'p-line below 8 8', 'p-line clear 29 29', 'p-line integer 41 41'])
    expect(table(shape, r => r.order)).toEqual(['asc 14 14', 'both 90 90', 'desc 10 10'])
  })
})
