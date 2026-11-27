// @vitest-environment node
/**
 * The across-midnight follow-up (brief cf-midnight, Phase 1 task 1): the survey (scripts/cf-refs/midnight-survey.ts),
 * pinned. Every path that can paint, cache or recalculate CF across midnight resolves to a line; on every surveyed case,
 * production's service one day after the load withholds every clock-rule cell, withholds nothing outside the clock and
 * reader cells, and changes nothing outside them.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { MIDNIGHT_PATHS, midnightCaseSurveyOf, resolveMidnightPaths, type MidnightCaseSurvey } from '../scripts/cf-refs/midnight-survey'

let survey: MidnightCaseSurvey[] = []
beforeAll(async () => { survey = await midnightCaseSurveyOf() }, 600_000)

describe('the paths', () => {
  it('every path resolves to a line of its file', () => {
    const paths = resolveMidnightPaths()
    expect(paths.length).toBe(MIDNIGHT_PATHS.length)
    for (const p of paths) expect(p.line, p.id).toBeGreaterThan(0)
    expect(new Set(paths.map(p => p.id)).size).toBe(paths.length)
  })
})

describe('the cells', () => {
  it('surveys every case with a workbook and rules', () => {
    expect(survey.length).toBe(90)
  })
  it('every clock-rule cell is withheld the next day, and nothing outside the clock and reader cells is withheld or changes', () => {
    for (const s of survey) {
      const reached = new Set([...s.clockCells, ...s.readerCells])
      const withheld = new Set(s.withheldNextDay)
      for (const r of s.clockCells) expect(withheld.has(r), `${s.id} ${r}: a clock cell is withheld`).toBe(true)
      for (const r of s.withheldNextDay) expect(reached.has(r), `${s.id} ${r}: withheld outside the clock and reader cells`).toBe(true)
      for (const r of s.changedNextDay) expect(reached.has(r), `${s.id} ${r}: changed outside the clock and reader cells`).toBe(true)
    }
  })
  it('the cases crossing midnight reaches, pinned', () => {
    const reached = Object.fromEntries(survey.filter(s => s.clockCells.length || s.readerCells.length || s.withheldNextDay.length || s.changedNextDay.length || s.taintedCells === 'all' || s.taintedCells.length)
      .map(s => [s.id, { clock: s.clockCells.length, tainted: s.taintedCells === 'all' ? 'all' : s.taintedCells.length, readers: s.readerCells.length, withheld: s.withheldNextDay.length, changed: s.changedNextDay.length }]))
    expect(reached).toEqual({
      'stage1-timeperiod-today-clock': { clock: 140, tainted: 20, readers: 140, withheld: 140, changed: 140 },
      'stage1-timeperiod-today-clock-2': { clock: 30, tainted: 10, readers: 30, withheld: 30, changed: 30 },
      'stage1-clock-across-midnight': { clock: 24, tainted: 22, readers: 10, withheld: 30, changed: 30 },
      'stage5-predicate-timeperiod': { clock: 0, tainted: 1, readers: 0, withheld: 0, changed: 0 },
      'stage5-clock-today': { clock: 9, tainted: 1, readers: 0, withheld: 9, changed: 9 },
      'stage5-calc-manual': { clock: 4, tainted: 1, readers: 0, withheld: 4, changed: 0 },
    })
  })
  it('the across-midnight case: the clock columns and the TODAY() cells under literal rules are withheld, the clock-free column is not', () => {
    const s = survey.find(x => x.id === 'stage1-clock-across-midnight')!
    const cols = (refs: readonly string[]) => [...new Set(refs.map(r => r.replace(/\d+/g, '')))].sort()
    expect(cols(s.withheldNextDay)).toEqual(['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'])
    expect(s.withheldNextDay.some(r => r.startsWith('J'))).toBe(false)
  })
})
