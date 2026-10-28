// @vitest-environment node
/**
 * Stage 1, 2 and 3 XLSX sparkline reference comparison (plan Tasks 7, 9 and 13, spec section 8 and section 11). Every case renders through
 * the viewer's static sheet renderer at the capture scale and is compared with Excel for Mac's copy-picture reference
 * under the tolerance in docs/validation/xlsx-sparklines.md (Cases table). Cases marked *unverified* there are skipped
 * and must carry a reason. Negative controls perturb the parsed model and must FAIL the same tolerances.
 */
import { copyFileSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { combine, compareCase, compareXlsxWriter, loadCaseSheet, main, readCaseTable, readSidecar, REFS, refBase, stageCases, UnusableReferenceError, withinTolerance, XLSXWRITER, XLSXWRITER_PARTS, type CaseEntry } from '../scripts/sparkline-compare'
import { captureRange } from '../scripts/sparkline-refs/cases'
import type { XlsxSheet } from '../src/xlsx/types'
import type { ResolvedSparkline } from '../src/xlsx/sparklines/resolve'

const table = readCaseTable()
/** Every Stage 1 case is expected to have references; none is filtered out by what happens to be on disk. */
const cases = stageCases(1)
/** Cases first captured for Stage 2 (column and win/loss bars). */
const stage2Cases = stageCases(2).filter(c => c.stage === 2)
/** Cases first captured for Stage 3 (date axis and right-to-left). */
const stage3Cases = stageCases(3).filter(c => c.stage === 3)
const byId = (id: string) => {
  const c = stageCases(3).find(x => x.id === id)
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

describe('XLSX sparkline references (Stage 3: date axis and right-to-left)', () => {
  it('there are 9 Stage 3 cases, each verified with a Cases row and complete references', () => {
    expect(stage3Cases.map(c => c.id)).toEqual(['date-uneven', 'date-unsorted', 'date-missing-range', 'date-length-mismatch', 'date-text', 'date-fractional', 'rtl-line', 'rtl-hl-fl', 'rtl-column'])
    for (const c of stage3Cases) {
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

  it('stage-3 references pass within tolerance', async () => {
    const failures: string[] = []
    for (const c of stage3Cases) {
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
    it('the date axis off: line, column and win/loss fall back to even slots', async () => {
      await fails('date-uneven', s => { delete inRow(s, 1).axis })
      await fails('date-uneven', s => { delete inRow(s, 2).axis })
      await fails('date-uneven', s => { delete inRow(s, 3).axis })
      await fails('date-length-mismatch', s => { delete inRow(s, 2).axis })
    }, 60_000)
    it('the control row (dates present, dateAxis off) must stay on even slots: giving it the date axis fails', async () => {
      await fails('date-uneven', s => { inRow(s, 4).axis = { ...inRow(s, 1).axis! } })
    }, 60_000)
    it('unsorted dates treated as data order (the date axis kept, the values not sorted by date)', async () => {
      const dataOrder = (s: XlsxSheet, row: number) => {
        const sp = inRow(s, row)
        sp.values = [4, 7, 3, 9, 2, 8, 5, 6]
        sp.highIndices = [3]; sp.lowIndices = [4]; sp.firstIndex = 0; sp.lastIndex = 7
      }
      await fails('date-unsorted', s => dataOrder(s, 1))
      await fails('date-unsorted', s => dataOrder(s, 2))
    }, 60_000)
    it('a dateAxis group with no date range drawn at even slots (the spec 5.4 hypothesis, refuted by Excel)', async () => {
      await fails('date-missing-range', s => {
        const group = s.sparklineGroups![0]
        const sp: ResolvedSparkline = { host: { row: 0, col: 9 }, dataSheet: 'Sheet1', sourceRows: Array(8).fill(0), group, values: [4, 7, 3, 9, 2, 8, 5, 6], bridged: [], firstIndex: 0, lastIndex: 7, highIndices: [3], lowIndices: [4], min: 2, max: 9 }
        s.resolvedSparklines = new Map([[0, new Map([[9, sp]])]])
      })
    }, 60_000)
    it('date cells that are text or empty kept as points: the dropped point put back', async () => {
      await fails('date-text', s => {
        const sp = inRow(s, 1)
        sp.values = [4, 7, 3, 9, 2, 8, 5, 6]
        sp.axis = { positions: [0, 1, 3, 10, 11, 20, 21, 30].map(o => (o + 0.5) / 31), slot: 1 / 31 }
      })
    }, 60_000)
    it('fractional date serials placed at time-of-day positions must fail the whole-day slot tolerance', async () => {
      await fails('date-fractional', s => {
        const dates = [45000, 45000.25, 45000.5, 45001.75, 45002, 45003.5, 45004.1, 45005]
        const slot = 1 / 6
        inRow(s, 2).axis = { positions: dates.map(date => (date - 45000 + 0.5) * slot), slot }
      })
    }, 60_000)
    it('right-to-left off: line, colours and bars unmirrored', async () => {
      await fails('rtl-hl-fl', s => { inRow(s, 1).group.rightToLeft = false })
      await fails('rtl-line', s => { inRow(s, 1).group.rightToLeft = false })
      await fails('rtl-line', s => { inRow(s, 2).group.rightToLeft = false })
      await fails('rtl-column', s => { inRow(s, 1).group.rightToLeft = false })
      await fails('rtl-column', s => { inRow(s, 2).group.rightToLeft = false })
    }, 60_000)
    it('right-to-left colours in display order instead of data order: first and last swapped', async () => {
      await fails('rtl-hl-fl', s => { const sp = inRow(s, 1); sp.firstIndex = 7; sp.lastIndex = 0 })
      await fails('rtl-column', s => { const sp = inRow(s, 1); sp.firstIndex = 7; sp.lastIndex = 0 })
    }, 60_000)
  })
})

describe('XLSX sparkline references (XlsxWriter-written workbook, an independent writer)', () => {
  const entries = XLSXWRITER_PARTS.map(p => ({ ...p, entry: table.get(p.id) as CaseEntry }))

  it('has Cases rows, a plain capture and a valid sidecar', () => {
    for (const p of entries) { expect(p.entry, p.id).toBeDefined(); expect(p.entry.tolerance, p.id).toBeDefined(); expect(p.entry.unverifiedReason, p.id).toBeUndefined() }
    expect(entries.flatMap(p => p.rows).sort((a, b) => a - b)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
    expect(existsSync(join(REFS, `${XLSXWRITER.id}.png`))).toBe(true)
    expect(readSidecar(join(REFS, `${XLSXWRITER.id}.json`), { sheet: XLSXWRITER.sheet, range: XLSXWRITER.range }).pxPerPt).toBe(1)
  })

  it('every host row passes within its Cases tolerance at the forced 204 x 60 cell', async () => {
    const failures: string[] = []
    for (const p of entries) {
      const r = await compareXlsxWriter({ rows: p.rows })
      expect(r.refSize).toEqual({ w: 204, h: 660 })
      if (!withinTolerance(r.diff, p.entry.tolerance!)) failures.push(`${p.id}: ${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`)
    }
    expect(failures).toEqual([])
  }, 60_000)

  it('an unsupported sidecar or a missing capture fails, never passes silently', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'sparkline-xw-'))
    temps.push(dir)
    await expect(compareXlsxWriter({ refsDir: dir })).rejects.toThrow(UnusableReferenceError)
    copyFileSync(join(REFS, `${XLSXWRITER.id}.png`), join(dir, `${XLSXWRITER.id}.png`))
    const sidecar = { ...JSON.parse(readFileSync(join(REFS, `${XLSXWRITER.id}.json`), 'utf8')), range: 'I1:J2' }
    writeFileSync(join(dir, `${XLSXWRITER.id}.json`), JSON.stringify(sidecar))
    await expect(compareXlsxWriter({ refsDir: dir })).rejects.toThrow(/range "I1:J2"/)
  })

  describe('negative controls fail', () => {
    const MAIN = XLSXWRITER_PARTS[0]
    const failsMain = async (mutate: (sheet: XlsxSheet) => void) => {
      const r = await compareXlsxWriter({ rows: MAIN.rows, mutate })
      expect(withinTolerance(r.diff, table.get(MAIN.id)!.tolerance!), `${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`).toBe(false)
    }
    it('the date axis off on J6, rightToLeft off on J4', async () => {
      await failsMain(s => { delete inRow(s, 6).axis })
      await failsMain(s => { inRow(s, 4).group.rightToLeft = false })
    }, 60_000)
    it('the win/loss axis off on J3, a custom-scale bar value on J7', async () => {
      await failsMain(s => { inRow(s, 3).group.displayXAxis = false })
      await failsMain(s => { (inRow(s, 7).values as number[])[2] = 90 })
    }, 60_000)
    it('the row-9 part, with its wide measured tolerance, still catches a gross vertex error', async () => {
      const part = XLSXWRITER_PARTS[1], tol = table.get(part.id)!.tolerance!
      const r = await compareXlsxWriter({ rows: part.rows, mutate: s => { (inRow(s, 9).values as number[])[3] = 40 } })
      expect(withinTolerance(r.diff, tol), `${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`).toBe(false)
    }, 60_000)
    it('the row-9 tolerance also rejects a mid-sized five-pixel vertex shift', async () => {
      const part = XLSXWRITER_PARTS[1], tol = table.get(part.id)!.tolerance!
      const r = await compareXlsxWriter({ rows: part.rows, mutate: s => { (inRow(s, 9).values as number[])[3] = 13 } })
      expect(withinTolerance(r.diff, tol), `${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`).toBe(false)
    }, 60_000)
  })
})

describe('XLSX sparkline colour fidelity (LibreOffice trap: negative tints)', () => {
  it('the default series colour (accent1, tint -0.5) resolves within 2 levels of Excel\'s captured (32,55,100)', async () => {
    const series = first(await loadCaseSheet(byId('line-basic'), 'Sheet1')).group.colors.series!
    const rgb = [2, 4, 6].map(i => parseInt(series.slice(i, i + 2), 16))
    ;[32, 55, 100].forEach((v, i) => expect(Math.abs(rgb[i] - v), `${series} channel ${i}`).toBeLessThanOrEqual(2))
  })
})
