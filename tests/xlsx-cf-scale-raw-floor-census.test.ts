// @vitest-environment node
/**
 * The raw-floor follow-up (brief cf-rawfloor), Phase 1 task 1: the census of W1 (docs/validation/xlsx-cf.md §14.z.17) over
 * every captured interior cell of a 2-colour scale or a 3-colour lower segment, Stage 2 and the Stage 5 painted colour-scale
 * cells (scripts/cf-refs/raw-floor-census.ts; docs/validation/xlsx-cf-recipes-stage2.md §11.1). Raw floor, floor(f·|d|)
 * per channel with f production's double, is Excel's fill at every one of them. This checks the census, not Excel. The census
 * as run before stage2-scale-raw-floor was captured is pinned first (recipe §11.1); with that capture (ce3f701d) included it
 * is pinned at the end (recipe §11.7): still exact.
 */
import { beforeAll, describe, expect, it, afterAll } from 'vitest'
import { asBeforePercentOrder, asRegistered } from './helpers/cf-percent-order-asof'
import { channelTable, censusTable, productClass, rawFloorCensusOf, type CensusRow } from '../scripts/cf-refs/raw-floor-census'
import { s2RawFloorCells } from '../scripts/cf-refs/cases'
import type { SegmentCollection } from '../scripts/cf-refs/num-segment-candidates'

// As of its capture: before the percent-order registration (§14.z.24), percent thresholds lerp-first and no
// scale:formula-bounds-v2 reading (tests/helpers/cf-percent-order-asof.ts).
beforeAll(asBeforePercentOrder)
afterAll(asRegistered)

const HEAVY = 300_000
let census: Awaited<ReturnType<typeof rawFloorCensusOf>>
let rows: CensusRow[] = []
let collection: SegmentCollection
beforeAll(async () => {
  // The census as run before stage2-scale-raw-floor was captured (recipe §11.1); the capture is pinned separately below.
  census = await rawFloorCensusOf(undefined, undefined, { include: 'pre-capture' })
  rows = census.rows
  collection = census.collection
}, HEAVY)
const short = (r: CensusRow): string => `${r.caseId.replace(/^stage\d-(scale-)?/, '')}:${r.ref}`

describe('the W1 census (every captured 2-colour or 3-colour lower-segment interior cell)', () => {
  it('1235 captured segment cells: 692 interior 2-colour / lower-segment cells, 268 endpoints or clamped (raw floor exact there too), 275 upper-segment; no harness fault', () => {
    expect(collection.cells.length).toBe(1235)
    expect(collection.faults).toEqual([])
    expect(rows.length).toBe(692)
    expect([census.endpoints, census.endpointsRawExact, census.upper]).toEqual([268, 268, 275])
  })

  it('W1 is exact: raw floor equals Excel\'s DisplayFormat fill at all 692; guarded truncation misses exactly the 14 below-integer cells', () => {
    expect(rows.filter(r => r.rawFloor !== r.excel)).toEqual([])
    expect(rows.filter(r => r.guarded !== r.excel).map(short)).toEqual(rows.filter(r => r.cls === 'below').map(short))
    expect(rows.filter(r => r.cls === 'below').map(short)).toEqual(['A3', 'A4', 'D2', 'D4', 'E7', 'E8', 'E9', 'E10', 'E12', 'E13', 'E14', 'H8', 'H18', 'I5'].map(x => `upper-window:${x}`))
    expect(rows.filter(r => r.cls === 'above').map(short)).toEqual(['num-segment:K15', 'upper-window:C12', 'upper-window:H19', 'upper-window:I2', 'cfvo-scope-v2:W17', 'cfvo-scope-v2:W20'])
  })

  it('by shape (scale, bounds, colour order, class): cells, raw-floor exact, guarded exact', () => {
    expect(censusTable(rows).map(t => `${t.key} ${t.cells} ${t.rawFloor} ${t.guarded}`)).toEqual([
      '2c formula both clear 51 51 51',
      '2c formula both integer 35 35 35',
      '2c literal asc clear 5 5 5',
      '2c literal both above 5 5 5',
      '2c literal both below 7 7 0',
      '2c literal both clear 141 141 141',
      '2c literal both integer 206 206 206',
      '2c mixed both clear 17 17 17',
      '2c mixed both integer 15 15 15',
      '2c statistical both below 3 3 0',
      '2c statistical both clear 39 39 39',
      '2c statistical both integer 53 53 53',
      '3c-lower formula asc clear 1 1 1',
      '3c-lower literal asc clear 8 8 8',
      '3c-lower literal both clear 26 26 26',
      '3c-lower literal both integer 32 32 32',
      '3c-lower mixed asc clear 4 4 4',
      '3c-lower mixed both below 2 2 0',
      '3c-lower mixed both clear 3 3 3',
      '3c-lower mixed both integer 1 1 1',
      '3c-lower statistical asc clear 8 8 8',
      '3c-lower statistical asc integer 1 1 1',
      '3c-lower statistical both above 1 1 1',
      '3c-lower statistical both below 2 2 0',
      '3c-lower statistical both clear 13 13 13',
      '3c-lower statistical both integer 13 13 13',
    ])
    // Production withholds 74 of the 692 (the 14 guard-boundary cells and 60 literal num-segment lower-segment cells); W1 holds there too.
    expect(rows.filter(r => r.status === 'withheld').length).toBe(74)
  })

  it('per channel: both truncation directions hold below-integer channels, all at the raw-floor step', () => {
    expect(channelTable(collection).filter(t => t.key.endsWith(' below')).map(t => `${t.key} ${t.channels} ${t.rawFloor} ${t.guarded}`)).toEqual([
      '2c asc below 13 13 0', '2c desc below 17 17 0', '3c-lower asc below 8 8 0', '3c-lower desc below 4 4 0',
    ])
    for (const t of channelTable(collection)) expect(t.rawFloor, t.key).toBe(t.channels)
    expect(channelTable(collection).reduce((n, t) => n + t.channels, 0)).toBe(2067)
  })

  it('not covered: 32 numeric cells Excel fills that another production gate withholds before interpolation (overlap, hidden row, merge, composition), all 2-colour', () => {
    const gated = collection.excluded.filter(e => e.group === 'b')
    expect(gated.length).toBe(32)
    expect(gated.filter(e => e.threeColour).length).toBe(0)
    expect([...new Set(gated.map(e => e.caseId))].sort()).toEqual(['stage2-scale-composition', 'stage2-scale-scope', 'stage5-composition'])
  })

  it('productClass: below, above, integer, clear', () => {
    expect([56.99999999999999, 57.00000000000001, 57, 56.5].map(productClass)).toEqual(['below', 'above', 'integer', 'clear'])
  })
})

describe('after the capture of stage2-scale-raw-floor (ce3f701d): the census on every capture (recipe §11.7)', () => {
  let post: Awaited<ReturnType<typeof rawFloorCensusOf>>
  beforeAll(async () => { post = await rawFloorCensusOf() }, HEAVY)
  const designated = s2RawFloorCells().filter(c => c.configuration !== undefined)

  it('1421 segment cells: 856 interior 2-colour / lower-segment cells (164 from the capture), 268 endpoints or clamped, 297 upper-segment; no harness fault', () => {
    expect(post.collection.cells.length).toBe(1421)
    expect(post.collection.faults).toEqual([])
    expect(post.rows.length).toBe(856)
    expect(post.rows.filter(r => r.caseId === 'stage2-scale-raw-floor').length).toBe(164)
    expect([post.endpoints, post.endpointsRawExact, post.upper]).toEqual([268, 268, 297])
  })

  it('W1 is still exact at all 856; guarded truncation misses exactly the 85 below-integer cells (ed0fccb6\'s 14 and this capture\'s 67 evidence and 4 literal cells)', () => {
    expect(post.rows.filter(r => r.rawFloor !== r.excel)).toEqual([])
    expect(post.rows.filter(r => r.guarded !== r.excel).map(short)).toEqual(post.rows.filter(r => r.cls === 'below').map(short))
    const evidence = designated.filter(c => c.role === 'evidence').map(c => `raw-floor:${c.ref}`)
    expect(post.rows.filter(r => r.cls === 'below').map(short).sort()).toEqual([...['A3', 'A4', 'D2', 'D4', 'E7', 'E8', 'E9', 'E10', 'E12', 'E13', 'E14', 'H8', 'H18', 'I5'].map(x => `upper-window:${x}`), ...evidence].sort())
    expect(evidence.length).toBe(71)
    // Above an integer: the 6 earlier cells and this capture's 9 lower-segment / 2-colour controls, all at k.
    expect(post.rows.filter(r => r.cls === 'above').map(short).sort()).toEqual(['num-segment:K15', 'upper-window:C12', 'upper-window:H19', 'upper-window:I2', 'cfvo-scope-v2:W17', 'cfvo-scope-v2:W20', ...['A8', 'A15', 'D7', 'D14', 'F5', 'F15', 'H7', 'I8', 'J6'].map(x => `raw-floor:${x}`)].sort())
  })

  it('by shape, and per channel: every row and every channel at raw floor; the below-integer channels in both directions', () => {
    for (const t of censusTable(post.rows)) expect(t.rawFloor, t.key).toBe(t.cells)
    expect(censusTable(post.rows).filter(t => t.key.endsWith(' below')).map(t => `${t.key} ${t.cells} ${t.rawFloor} ${t.guarded}`)).toEqual([
      '2c literal both below 29 29 0', '2c statistical both below 24 24 0', '3c-lower literal both below 4 4 0', '3c-lower mixed both below 13 13 0', '3c-lower statistical both below 15 15 0',
    ])
    expect(channelTable(post.collection).filter(t => t.key.endsWith(' below')).map(t => `${t.key} ${t.channels} ${t.rawFloor} ${t.guarded}`)).toEqual([
      '2c asc below 72 72 0', '2c desc below 87 87 0', '3c-lower asc below 59 59 0', '3c-lower desc below 37 37 0',
    ])
    for (const t of channelTable(post.collection)) expect(t.rawFloor, t.key).toBe(t.channels)
    expect(channelTable(post.collection).reduce((n, t) => n + t.channels, 0)).toBe(2559)
    // The harness's status (as when the searches ran): the 74 earlier, this capture's 67 guard-boundary cells (painted at
    // raw floor since the registration) and its 4 literal num-segment cells.
    expect(post.rows.filter(r => r.status === 'withheld').length).toBe(74 + 67 + 4)
    expect(post.collection.excluded.filter(e => e.group === 'b').length).toBe(32)
  })
})
