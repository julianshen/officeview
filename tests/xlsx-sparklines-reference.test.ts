// @vitest-environment node
/**
 * Stage 1 XLSX sparkline reference comparison (plan Task 7, spec section 8 and section 11). Every case renders through
 * the viewer's static sheet renderer at the capture scale and is compared with Excel for Mac's copy-picture reference
 * under the tolerance in docs/validation/xlsx-sparklines.md (Cases table). Cases marked *unverified* there are skipped
 * and must carry a reason. Negative controls perturb the parsed model and must FAIL the same tolerances.
 */
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { combine, compareCase, readCaseTable, REFS, refBase, stageCases, withinTolerance, type CaseEntry } from '../scripts/sparkline-compare'
import type { XlsxSheet } from '../src/xlsx/types'
import type { ResolvedSparkline } from '../src/xlsx/sparklines/resolve'

const table = readCaseTable()
const cases = stageCases(1).filter(c => c.sheets.some(s => s.groups.length && existsSync(join(REFS, `${refBase(c, s.name, false)}.png`))))
const byId = (id: string) => {
  const c = stageCases(1).find(x => x.id === id)
  if (!c) throw new Error(`no case ${id}`)
  return c
}
const sparklines = (sheet: XlsxSheet): ResolvedSparkline[] => [...(sheet.resolvedSparklines ?? new Map()).values()].flatMap(m => [...m.values()]) as ResolvedSparkline[]
const first = (sheet: XlsxSheet): ResolvedSparkline => sparklines(sheet)[0]
const groups = (sheet: XlsxSheet) => [...new Set(sparklines(sheet).map(s => s.group))]
/** Plot height in capture px for the 60-pt host rows (59-pt inner box minus the insets 4.08 + 3.06 pt). */
const PLOT_H = 59 - 4.08 - 3.06

describe('XLSX sparkline references (Stage 1)', () => {
  it('every Stage 1 reference has a Cases row with a tolerance or an unverified reason', () => {
    expect(cases.length).toBe(30)
    for (const c of cases) {
      const entry = table.get(c.id)
      expect(entry, c.id).toBeDefined()
      if (entry!.unverifiedReason !== undefined) expect(entry!.unverifiedReason.length, c.id).toBeGreaterThan(20)
      else expect(entry!.tolerance, c.id).toBeDefined()
    }
  })

  it('stage-1 references pass within tolerance', async () => {
    const failures: string[] = []
    for (const c of cases) {
      const entry = table.get(c.id) as CaseEntry
      if (entry.unverifiedReason !== undefined) continue
      const d = combine(await compareCase(c))
      if (!withinTolerance(d, entry.tolerance!)) failures.push(`${c.id}: ${d.meanAbs.toFixed(4)} ${d.rowMax.toFixed(4)} strong ${d.strong}`)
    }
    expect(failures).toEqual([])
  }, 120_000)

  describe('negative controls fail', () => {
    const fails = async (id: string, mutate: (sheet: XlsxSheet) => void) => {
      const d = combine(await compareCase(byId(id), { mutate }))
      expect(withinTolerance(d, table.get(id)!.tolerance!), `${id}: ${d.meanAbs.toFixed(4)} ${d.rowMax.toFixed(4)} strong ${d.strong}`).toBe(false)
    }
    // line-basic D = [4,7,3,9,2,8,5,6] on [2, 9]: one value unit is PLOT_H / 7 px. idx 2 and 5 match the reference
    // exactly, so moving them 2 px is a 2-px vertex error.
    it('a vertex 2 px off', async () => {
      await fails('line-basic', s => { (first(s).values as number[])[2] += 2 * 7 / PLOT_H })
      await fails('line-basic', s => { (first(s).values as number[])[5] -= 2 * 7 / PLOT_H })
      await fails('line-merged-host', s => { (first(s).values as number[])[3] -= 2 * 7 / PLOT_H })
    }, 60_000)
    it('a wrong series colour (tint −0.25 instead of −0.5, and no tint)', async () => {
      await fails('line-basic', s => { first(s).group.colors.series = 'FF2F5597' })
      await fails('line-basic', s => { first(s).group.colors.series = 'FF4472C4' })
    }, 60_000)
    it('missing markers, including a single missing special point', async () => {
      await fails('line-markers', s => { first(s).group.markers = false })
      await fails('line-points-hl-fl', s => { first(s).group.high = false })
      await fails('line-points-hl-fl', s => { first(s).group.last = false })
    }, 60_000)
    it('the axis one row off and swapped point colours', async () => {
      // A slightly lower scale minimum moves y(0) on [−4, 4] from row 31 to row 30.
      await fails('line-axis', s => { for (const sp of sparklines(s)) sp.min -= 0.1 })
      await fails('line-precedence', s => { for (const g of groups(s)) [g.colors.high, g.colors.low] = [g.colors.low, g.colors.high] })
    }, 60_000)
  })
})
