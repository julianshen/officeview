// @vitest-environment node
/**
 * Stage 1 and Stage 2 XLSX sparkline reference comparison (plan Tasks 7 and 9, spec section 8 and section 11). Every case renders through
 * the viewer's static sheet renderer at the capture scale and is compared with Excel for Mac's copy-picture reference
 * under the tolerance in docs/validation/xlsx-sparklines.md (Cases table). Cases marked *unverified* there are skipped
 * and must carry a reason. Negative controls perturb the parsed model and must FAIL the same tolerances.
 */
import { copyFileSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { combine, compareCase, main, readCaseTable, readSidecar, REFS, refBase, stageCases, UnusableReferenceError, withinTolerance, type CaseEntry } from '../scripts/sparkline-compare'
import { captureRange } from '../scripts/sparkline-refs/cases'
import type { XlsxSheet } from '../src/xlsx/types'
import type { ResolvedSparkline } from '../src/xlsx/sparklines/resolve'

const table = readCaseTable()
/** Every Stage 1 case is expected to have references; none is filtered out by what happens to be on disk. */
const cases = stageCases(1)
/** Cases first captured for Stage 2 (column and win/loss bars). */
const stage2Cases = stageCases(2).filter(c => c.stage === 2)
const byId = (id: string) => {
  const c = stageCases(2).find(x => x.id === id)
  if (!c) throw new Error(`no case ${id}`)
  return c
}
const sparklines = (sheet: XlsxSheet): ResolvedSparkline[] => [...(sheet.resolvedSparklines ?? new Map()).values()].flatMap(m => [...m.values()]) as ResolvedSparkline[]
const first = (sheet: XlsxSheet): ResolvedSparkline => sparklines(sheet)[0]
/** The sparkline hosted in the 1-based sheet row `row` (Stage 2 cases put one sparkline per row, column J). */
const inRow = (sheet: XlsxSheet, row: number): ResolvedSparkline => {
  const found = sparklines(sheet).find(s => s.host.row === row - 1)
  if (!found) throw new Error(`no sparkline in row ${row}`)
  return found
}
const groups = (sheet: XlsxSheet) => [...new Set(sparklines(sheet).map(s => s.group))]
/** Plot height in capture px for the 60-pt host rows (59-pt inner box minus the insets 4.08 + 3.06 pt). */
const PLOT_H = 59 - 4.08 - 3.06

/** A temporary copy of line-basic's four reference files. */
const temps: string[] = []
afterEach(() => { for (const dir of temps.splice(0)) rmSync(dir, { recursive: true, force: true }) })
const copyRefs = (id: string): string => {
  const dir = mkdtempSync(join(tmpdir(), 'sparkline-refs-'))
  temps.push(dir)
  for (const f of readdirSync(REFS)) if (f.startsWith(`${id}.`) || f.startsWith(`${id}-bordered.`)) copyFileSync(join(REFS, f), join(dir, f))
  return dir
}

describe('XLSX sparkline references (Stage 1)', () => {
  it('every Stage 1 case has a Cases row, and every verified case both images and both valid sidecars', () => {
    expect(cases.length).toBe(30)
    for (const c of cases) {
      const entry = table.get(c.id)
      expect(entry, c.id).toBeDefined()
      if (entry!.unverifiedReason !== undefined) { expect(entry!.unverifiedReason.length, c.id).toBeGreaterThan(20); continue }
      expect(entry!.tolerance, c.id).toBeDefined()
      const sheets = c.sheets.filter(s => captureRange(s))
      expect(sheets.length, c.id).toBeGreaterThan(0)
      for (const s of sheets) for (const bordered of [false, true]) {
        const base = join(REFS, refBase(c, s.name, bordered))
        expect(existsSync(`${base}.png`), `${base}.png`).toBe(true)
        expect(readSidecar(`${base}.json`, { sheet: s.name, range: captureRange(s)! }).pxPerPt, `${base}.json`).toBe(1)
      }
    }
  })

  it('stage-1 references pass within tolerance', async () => {
    const failures: string[] = []
    for (const c of cases) {
      const entry = table.get(c.id) as CaseEntry
      if (entry.unverifiedReason !== undefined) continue
      try {
        const d = combine(await compareCase(c))
        if (!withinTolerance(d, entry.tolerance!)) failures.push(`${c.id}: ${d.meanAbs.toFixed(4)} ${d.rowMax.toFixed(4)} strong ${d.strong}`)
      } catch (error) { failures.push(`${c.id}: ${(error as Error).message}`) }
    }
    expect(failures).toEqual([])
  }, 120_000)

  describe('unusable references fail, never pass silently', () => {
    const quiet = () => { vi.spyOn(console, 'log').mockImplementation(() => {}); vi.spyOn(console, 'error').mockImplementation(() => {}) }
    afterEach(() => { vi.restoreAllMocks() })

    it('the temporary copy itself compares and passes (control for the cases below)', async () => {
      quiet()
      expect(await main(['--case', 'line-basic', '--refs', copyRefs('line-basic')])).toBe(0)
    })
    it('a missing -bordered twin fails compareCase and the CLI', async () => {
      const dir = copyRefs('line-basic')
      rmSync(join(dir, 'line-basic-bordered.png'))
      await expect(compareCase(byId('line-basic'), { refsDir: dir })).rejects.toThrow(UnusableReferenceError)
      await expect(compareCase(byId('line-basic'), { refsDir: dir })).rejects.toThrow(/missing reference .*line-basic-bordered\.png/)
      quiet()
      expect(await main(['--case', 'line-basic', '--refs', dir])).toBe(1)
    })
    it('a missing sidecar fails', async () => {
      const dir = copyRefs('line-basic')
      rmSync(join(dir, 'line-basic.json'))
      await expect(compareCase(byId('line-basic'), { refsDir: dir })).rejects.toThrow(/missing sidecar/)
      quiet()
      expect(await main(['--case', 'line-basic', '--refs', dir])).toBe(1)
    })
    it('a sidecar with an unsupported scale, route or range fails', async () => {
      const edit = (dir: string, change: Record<string, unknown>) => {
        const path = join(dir, 'line-basic-bordered.json')
        writeFileSync(path, JSON.stringify({ ...JSON.parse(readFileSync(path, 'utf8')), ...change }))
      }
      for (const [change, message] of [
        [{ scale: 2, pxPerPt: 2 }, /pxPerPt 2 \(supported: 1/],
        [{ scale: 2 }, /scale 2 differs from pxPerPt 1/],
        [{ route: 'screenshot' }, /route "screenshot"/],
        [{ range: 'I1:J2' }, /range "I1:J2"/],
      ] as const) {
        const dir = copyRefs('line-basic')
        edit(dir, change)
        await expect(compareCase(byId('line-basic'), { refsDir: dir })).rejects.toThrow(message)
        quiet()
        expect(await main(['--case', 'line-basic', '--refs', dir])).toBe(1)
        vi.restoreAllMocks()
      }
    })
  })

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

describe('XLSX sparkline references (Stage 2: column and win/loss)', () => {
  it('there are 11 Stage 2 cases, each verified with a Cases row and complete references', () => {
    expect(stage2Cases.map(c => c.id)).toEqual(['col-basic', 'col-negative', 'col-scale-no-zero', 'col-hl-fl', 'col-empty-span', 'col-merged-host', 'col-tiny', 'wl-basic', 'wl-tiny-values', 'wl-zero', 'wl-hl-fl'])
    for (const c of stage2Cases) {
      const entry = table.get(c.id)
      expect(entry, c.id).toBeDefined()
      expect(entry!.unverifiedReason, c.id).toBeUndefined()
      expect(entry!.tolerance, c.id).toBeDefined()
      for (const s of c.sheets.filter(x => captureRange(x))) for (const bordered of [false, true]) {
        const base = join(REFS, refBase(c, s.name, bordered))
        expect(existsSync(`${base}.png`), `${base}.png`).toBe(true)
        expect(readSidecar(`${base}.json`, { sheet: s.name, range: captureRange(s)! }).pxPerPt, `${base}.json`).toBe(1)
      }
    }
  })

  it('stage-2 references pass within tolerance', async () => {
    const failures: string[] = []
    for (const c of stage2Cases) {
      const entry = table.get(c.id) as CaseEntry
      try {
        const d = combine(await compareCase(c))
        if (!withinTolerance(d, entry.tolerance!)) failures.push(`${c.id}: ${d.meanAbs.toFixed(4)} ${d.rowMax.toFixed(4)} strong ${d.strong}`)
      } catch (error) { failures.push(`${c.id}: ${(error as Error).message}`) }
    }
    expect(failures).toEqual([])
  }, 120_000)

  describe('negative controls fail', () => {
    const fails = async (id: string, mutate: (sheet: XlsxSheet) => void) => {
      const d = combine(await compareCase(byId(id), { mutate }))
      expect(withinTolerance(d, table.get(id)!.tolerance!), `${id}: ${d.meanAbs.toFixed(4)} ${d.rowMax.toFixed(4)} strong ${d.strong}`).toBe(false)
    }
    it('a bar whose value moves it, and a column filled where Excel leaves a gap', async () => {
      await fails('col-basic', s => { (inRow(s, 1).values as number[])[3] = 5 })
      await fails('col-merged-host', s => { (first(s).values as number[])[0] = 9 })
      await fails('col-empty-span', s => { (inRow(s, 1).values as Array<number | null>)[1] = 5 })
    }, 60_000)
    it('a shifted scale and a wrong baseline', async () => {
      await fails('col-scale-no-zero', s => { inRow(s, 1).min -= 5 })
      await fails('col-negative', s => { inRow(s, 1).max += 3 })
    }, 60_000)
    it('negative colouring and the axis, columns and win/loss', async () => {
      await fails('col-negative', s => { inRow(s, 1).group.negative = false })
      await fails('col-negative', s => { inRow(s, 4).group.displayXAxis = false })
      await fails('wl-basic', s => { inRow(s, 2).group.negative = false })
      await fails('wl-basic', s => { inRow(s, 3).group.displayXAxis = false })
      await fails('wl-zero', s => { for (const sp of sparklines(s)) sp.group.displayXAxis = !sp.group.displayXAxis })
    }, 60_000)
    it('a win bar turned into a loss bar', async () => {
      await fails('wl-basic', s => { const v = inRow(s, 1).values as number[]; v[0] = -Math.abs(v[0]) || -1; v[1] = -Math.abs(v[1]) })
    }, 60_000)
    it('missing or swapped element colours', async () => {
      await fails('col-hl-fl', s => { inRow(s, 1).group.high = false })
      await fails('col-hl-fl', s => { inRow(s, 1).group.last = false })
      await fails('col-hl-fl', s => { const g = inRow(s, 3).group; [g.colors.high, g.colors.low] = [g.colors.low, g.colors.high] })
      await fails('wl-hl-fl', s => { inRow(s, 1).group.low = false })
    }, 60_000)
  })
})
