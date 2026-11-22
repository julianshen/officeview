// @vitest-environment node
/**
 * The upper-segment window follow-up (brief cf-upperwindow), Phase 1 task 1: the census of the captured cells inside the
 * 1e-11 window (not IEEE integers) and the candidate search of scripts/cf-refs/upper-window-candidates.ts, pinned to the
 * numbers of docs/validation/xlsx-cf-recipes-stage2.md §10.1. This checks the search, not Excel.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { evaluateUpperWindow, integerShape, scoredCells, upperWindowCandidates, upperWindowCells, windowCensus, windowShape, type UpperWindowResult } from '../scripts/cf-refs/upper-window-candidates'
import type { SegmentCell } from '../scripts/cf-refs/num-segment-candidates'

const HEAVY = 300_000
let cells: SegmentCell[] = []
let results: UpperWindowResult[] = []
beforeAll(async () => {
  cells = await upperWindowCells()
  results = evaluateUpperWindow(cells)
}, HEAVY)
const short = (k: string): string => k.replace(/^stage\d-(scale-)?/, '')

describe('the window census (every captured cell inside the 1e-11 window, not an IEEE integer)', () => {
  it('1101 cells; 155 in the integer shape (not scored); 946 scored', () => {
    expect(cells.length).toBe(1101)
    expect(cells.filter(c => integerShape(c.segment)).length).toBe(155)
    expect(scoredCells(cells).length).toBe(946)
  })

  it('13 cells: 10 upper-segment 3-colour (2 statistical-bound, both at guarded truncation), 3 two-colour (above an integer, at truncation)', () => {
    const rows = windowCensus(cells)
    expect(rows.map(r => `${short(r.caseId)}:${r.ref} ${r.colours}/${r.segment} ${r.bounds} ${r.outcome}`)).toEqual([
      'thresholds:E4 3/1 percentile/max guarded',
      'num-segment:K15 2/0 num/num both',
      'upper-whole:D19 3/1 percentile/max guarded',
      'upper-whole:F8 3/1 num/num guarded',
      'upper-whole:F11 3/1 num/num both',
      'upper-whole:F12 3/1 num/num raw-floor',
      'upper-whole:F13 3/1 num/num both',
      'upper-whole:F15 3/1 num/num raw-floor',
      'upper-whole:F16 3/1 num/num raw-floor',
      'upper-whole:F18 3/1 num/num raw-floor',
      'upper-whole:F19 3/1 num/num neither',
      'cfvo-scope-v2:W17 2/0 num/num both',
      'cfvo-scope-v2:W20 2/0 num/num both',
    ])
    // The window shape (what a window reading paints) is the ten upper-segment cells.
    expect(cells.filter(c => windowShape(c.segment)).map(c => c.ref)).toEqual(['E4', 'D19', 'F8', 'F11', 'F12', 'F13', 'F15', 'F16', 'F18', 'F19'])
    // Below an integer: E4, D19, F8 (4/39) at guarded truncation; F12, F15, F16, F18 one smaller. Above: F11, F13 at k; F19 one smaller.
    const cls = new Map(rows.map(r => [r.ref, r.channels.map(x => x.cls)]))
    for (const ref of ['E4', 'D19', 'F8']) expect(cls.get(ref), ref).toEqual(['below', 'none', 'none'])
    for (const ref of ['F12', 'F15', 'F16', 'F18']) expect(cls.get(ref), ref).toEqual(['below', 'below', 'below'])
    for (const ref of ['F13', 'F19']) expect(cls.get(ref), ref).toEqual(['above', 'above', 'above'])
  })
})

describe('the search: no survivor at any scope', () => {
  it('48603 candidates x 4 scopes; 0 survivors', () => {
    expect(upperWindowCandidates().length).toBe(48603)
    expect(results.length).toBe(48603 * 4)
    expect(results.filter(r => r.mismatches.length === 0)).toEqual([])
  })

  it('the best of each family at the window scope (10 cells)', () => {
    const best = (fam: string): number => Math.min(...results.filter(r => r.family === fam && r.scope === 'window').map(r => r.mismatches.length))
    expect(Object.fromEntries(['origin', 'f32', 'precision', 'rational', 'form', 'threshold', 'reading'].map(f => [f, best(f)]))).toEqual({ origin: 1, f32: 1, precision: 1, rational: 3, form: 1, threshold: 2, reading: 4 })
    // Every 1-mismatch candidate misses one literal-num cell (F19, F13, F12 or F16); none misses E4 or D19.
    const near = results.filter(r => r.scope === 'window' && r.mismatches.length === 1)
    expect(near.length).toBe(1199)
    expect([...new Set(near.map(r => short(r.mismatches[0]!)))].sort()).toEqual(['upper-whole:F12', 'upper-whole:F13', 'upper-whole:F16', 'upper-whole:F19'])
  })

  it('the two pre-registered readings: guarded misses F12, F15, F16, F18, F19; raw floor misses E4, D19, F8, F19', () => {
    const at = (id: string): string[] => results.find(r => r.id === id && r.scope === 'window')!.mismatches.map(short)
    expect(at('reading/guarded')).toEqual(['upper-whole:F12', 'upper-whole:F15', 'upper-whole:F16', 'upper-whole:F18', 'upper-whole:F19'])
    expect(at('reading/raw-floor')).toEqual(['thresholds:E4', 'upper-whole:D19', 'upper-whole:F8', 'upper-whole:F19'])
  })
})
