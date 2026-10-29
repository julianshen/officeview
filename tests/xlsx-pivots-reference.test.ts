// @vitest-environment node
/**
 * Stage 1 XLSX pivot, table, slicer and timeline reference comparison (plan 2026-10-29-xlsx-pivots Task 14, spec §9).
 * Every case renders through the interactive viewport painter at 75 % and is compared with Excel for Mac's capture
 * under the tolerance in docs/validation/xlsx-pivots.md (Cases table); scripts/pivot-compare.ts holds the method.
 * Cases marked *unverified* there are skipped and must carry a reason. Negative controls perturb the parsed model and
 * must FAIL the same tolerances, so the masks are shown not to hide what the comparison is for.
 *
 * The pixel comparison runs on macOS only, like the slideshow goldens: the masks follow the cell text the host's font
 * stack lays out (spills and indents), and the tolerances were fitted on the Mac that holds the references. A Linux
 * run of the pixel block is unverified. The structure, unusable-reference and corpus blocks run everywhere.
 */
import { copyFileSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { rangeRect } from '../src/xlsx/regionGrid'
import type { XlsxSheet } from '../src/xlsx/types'
import { compareCase, main, PIVOT_CASES, readPivotCaseTable, readSidecar, REFS, stageCases, UnusableReferenceError, type CompareOptions, type PivotRefCase } from '../scripts/pivot-compare'
import { withinTolerance, type Tolerance } from '../scripts/sparkline-compare'

const table = readPivotCaseTable()
const cases = stageCases(1)
const byId = (id: string): PivotRefCase => {
  const c = PIVOT_CASES.find(x => x.id === id)
  if (!c) throw new Error(`no case ${id}`)
  return c
}
const tolerance = (id: string): Tolerance => table.get(id)!.tolerance!

const temps: string[] = []
afterEach(() => { for (const dir of temps.splice(0)) rmSync(dir, { recursive: true, force: true }) })
/** A temporary copy of one case's reference image and sidecar. */
const copyRefs = (id: string): string => {
  const dir = mkdtempSync(join(tmpdir(), 'pivot-refs-'))
  temps.push(dir)
  for (const f of readdirSync(REFS)) if (f === `${id}.png` || f === `${id}.json`) copyFileSync(join(REFS, f), join(dir, f))
  return dir
}
const quiet = () => { vi.spyOn(console, 'log').mockImplementation(() => {}); vi.spyOn(console, 'error').mockImplementation(() => {}) }

describe('XLSX pivot references: cases and sidecars', () => {
  it('every Stage 1 case has a Cases row; every verified case a tolerance, a reference and a valid sidecar', () => {
    // 70 tables, 96 pivots, 17 slicers, 48 timelines, 5 measurement captures and 3 window screenshots.
    expect(cases.length).toBe(239)
    expect(new Set(cases.map(c => c.id)).size).toBe(cases.length)
    for (const c of cases) {
      const entry = table.get(c.id)
      expect(entry, c.id).toBeDefined()
      if (entry!.unverifiedReason !== undefined) { expect(entry!.unverifiedReason.length, c.id).toBeGreaterThan(20); continue }
      expect(entry!.status, c.id).toBe('pass')
      expect(entry!.tolerance, c.id).toBeDefined()
      expect(existsSync(join(REFS, `${c.id}.png`)), c.id).toBe(true)
      expect(readSidecar(join(REFS, `${c.id}.json`), c).pxPerPt, c.id).toBe(1)
    }
    // Only these two are unverified (their reasons are in the Cases notes).
    expect(cases.filter(c => table.get(c.id)?.unverifiedReason !== undefined).map(c => c.id).sort()).toEqual(['measure-apply-flags', 'pivot-chart-buttons-all-window'])
  })

  it('the Cases table lists no case the script does not know', () => {
    for (const id of table.keys()) expect(PIVOT_CASES.some(c => c.id === id), id).toBe(true)
  })
})

describe('unusable references fail, never pass silently', () => {
  afterEach(() => { vi.restoreAllMocks() })

  it.runIf(process.platform === 'darwin')('the temporary copy itself compares and passes (control for the cases below)', async () => {
    quiet()
    expect(await main(['--case', 'pivot-light-16', '--refs', copyRefs('pivot-light-16')])).toBe(0)
  })
  it('a missing reference image fails compareCase and the CLI', async () => {
    const dir = copyRefs('pivot-light-16')
    rmSync(join(dir, 'pivot-light-16.png'))
    await expect(compareCase(byId('pivot-light-16'), { refsDir: dir })).rejects.toThrow(/missing reference .*pivot-light-16\.png/)
    quiet()
    expect(await main(['--case', 'pivot-light-16', '--refs', dir])).toBe(1)
  })
  it('a missing sidecar fails', async () => {
    const dir = copyRefs('pivot-light-16')
    rmSync(join(dir, 'pivot-light-16.json'))
    await expect(compareCase(byId('pivot-light-16'), { refsDir: dir })).rejects.toThrow(/missing sidecar/)
  })
  it('a sidecar with another route, density, appearance, sheet, range or source fails', async () => {
    for (const [change, message] of [
      [{ route: 'window-screenshot' }, /route "window-screenshot"/],
      [{ pxPerPt: 2 }, /pxPerPt 2/],
      [{ appearance: 'printer' }, /appearance "printer"/],
      [{ sheet: 'Data' }, /sheet "Data"/],
      [{ range: 'not a range' }, /range "not a range"/],
      [{ source: 'pivot-dark-1.xlsx' }, /source "pivot-dark-1.xlsx"/],
    ] as const) {
      const dir = copyRefs('pivot-light-16'), path = join(dir, 'pivot-light-16.json')
      writeFileSync(path, JSON.stringify({ ...JSON.parse(readFileSync(path, 'utf8')), ...change }))
      await expect(compareCase(byId('pivot-light-16'), { refsDir: dir })).rejects.toThrow(UnusableReferenceError)
      await expect(compareCase(byId('pivot-light-16'), { refsDir: dir })).rejects.toThrow(message)
    }
  })
})

describe.skipIf(process.platform !== 'darwin')('XLSX pivot references (Stage 1, macOS)', () => {
  it('every verified case passes within its tolerance', async () => {
    const failures: string[] = []
    for (const c of cases) {
      const entry = table.get(c.id)!
      if (entry.unverifiedReason !== undefined) continue
      try {
        const r = await compareCase(c)
        if (!withinTolerance(r.diff, entry.tolerance!)) failures.push(`${c.id}: ${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`)
      } catch (error) { failures.push(`${c.id}: ${(error as Error).message}`) }
    }
    expect(failures).toEqual([])
  }, 180_000)

  describe('negative controls fail', () => {
    /** The case perturbed by `mutate` must fall outside its own tolerance. */
    const fails = async (id: string, mutate: CompareOptions['mutate']) => {
      const r = await compareCase(byId(id), { mutate })
      expect(withinTolerance(r.diff, tolerance(id)), `${id}: ${r.diff.meanAbs.toFixed(4)} ${r.diff.rowMax.toFixed(4)} strong ${r.diff.strong}`).toBe(false)
    }
    const restyle = (sheet: XlsxSheet) => { sheet.regionStyles = undefined; sheet.regionStylesResolved = false }

    it('another pivot style (PivotStyleLight16 drawn as PivotStyleDark1)', () => fails('pivot-light-16', sheet => {
      sheet.pivotTables![0]!.styleInfo = { ...sheet.pivotTables![0]!.styleInfo!, name: 'PivotStyleDark1' }
      restyle(sheet)
    }))
    it('no ± buttons on the pivot (showDrill off)', () => fails('pivot-medium-9', sheet => { sheet.pivotTables![0]!.showDrill = false }))
    it('table row stripes off (TableStyleMedium2)', () => fails('table-medium-2', sheet => {
      sheet.tables![0]!.styleInfo = { ...sheet.tables![0]!.styleInfo!, showRowStripes: false }
      restyle(sheet)
    }))
    it('a slicer item selected that Excel shows unselected (Dates)', () => fails('slicer-light1', sheet => {
      const content = sheet.drawings!.find(d => d.content?.kind === 'slicer')!.content!
      if (content.kind === 'slicer') content.slicer.items[1]!.selected = true
    }))
    it('the timeline selection one month later (April–June)', () => fails('measure-timeline-selection', sheet => {
      const content = sheet.drawings!.find(d => d.content?.kind === 'timeline')!.content!
      if (content.kind === 'timeline') content.timeline.selection = { start: Date.UTC(2024, 3, 1), end: Date.UTC(2024, 5, 30) }
    }))
    it('dropdowns missing from a window screenshot (showDataDropDown off)', () => fails('measure-field-buttons', sheet => { sheet.pivotTables![0]!.showDataDropDown = false }))
  })
})

// ---------------------------------------------------------------------------------------------------------------
// Corpus (step 3): every corpus pivot and table workbook parses, styles and paints without throwing, and the saved
// cells inside each pivot location and table are shown.
// ---------------------------------------------------------------------------------------------------------------

/** The corpus files of docs/validation/xlsx-pivots.md "Corpus (Task 1)". */
const CORPUS = [
  'poi-ExcelPivotTableSample.xlsx', 'poi-ExcelTables.xlsx', 'poi-tableStyle.xlsx',
  'lo-calc-Pivot2_Row_Compact.xlsx', 'lo-calc-Pivot2_Row_Grand_Subtotals.xlsx', 'lo-calc-Pivot3_Column_Grand_Subtotals.xlsx',
  'lo-calc-Pivot4_Column_Grand_Subtotals_SortDescending.xlsx', 'lo-calc-pivottable_outline_mode.xlsx', 'lo-calc-pivottable_tabular_mode.xlsx',
  'lo-calc-pivottable_rowcolpage_field_filter.xlsx', 'lo-calc-pivottable_date_field_filter.xlsx', 'lo-calc-pivot_dark1.xlsx',
  'lo-calc-pivotcompact.xlsx', 'lo-calc-two-data-fields.xlsx', 'lo-calc-calcfields.xlsx', 'lo-calc-shared-dategroup.xlsx',
  'lo-calc-shared-numgroup.xlsx', 'lo-calc-PivotTableFormatsSubTotals.xlsx', 'lo-calc-PivotTableCellFormatsTest_7_TwoRowTwoColumnFields_DataColor.xlsx',
  'lo-calc-PivotTableWithNoSourceData.xlsx', 'lo-calc-PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx',
  'lo-calc-TableStyleTest.xlsx', 'lo-calc-tableStyleInnerBorders.xlsx', 'lo-calc-tdf162963_TableWithTotalsEnabled.xlsx',
] as const
const CORPUS_DIR = resolve(process.cwd(), 'corpus')

describe('corpus pivots and tables: no throw, saved cells shown', () => {
  it('lists the 24 corpus files of the validation doc', () => {
    const doc = readFileSync(resolve(process.cwd(), 'docs/validation/xlsx-pivots.md'), 'utf8')
    expect(CORPUS.length).toBe(24)
    for (const name of CORPUS) expect(doc.includes(name.replace(/^lo-calc-|^poi-/, '')), name).toBe(true)
  })

  it.each(CORPUS.map(name => [name]))('%s', async (name) => {
    const path = join(CORPUS_DIR, name)
    if (!existsSync(path)) return // the corpus is fetched (bun scripts/fetch-corpus.ts), as CI does
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(path)))
    await ensureTableStyleEngine(doc.sheets)
    let regions = 0
    for (const sheet of doc.sheets) {
      const rects = [
        ...(sheet.pivotTables ?? []).map(def => rangeRect(def.location.ref)),
        ...(sheet.tables ?? []).map(t => ({ minRow: t.extent.firstRow, minCol: t.extent.firstCol, maxRow: t.extent.firstRow + t.extent.rows - 1, maxCol: t.extent.firstCol + t.extent.cols - 1 })),
      ].filter(rect => rect !== undefined)
      const geometry = buildSheetGeometry(sheet)
      expect(geometry.kind, `${name} ${sheet.name}`).toBe('ready')
      if (geometry.kind !== 'ready') continue
      for (const rect of rects) {
        regions++
        const { cols, rows, order } = geometry.metrics
        const x0 = cols.start(rect.minCol), y0 = rows.start(order.sourceToDisplay(rect.minRow))
        const width = Math.min(4000, cols.start(rect.maxCol + 1) - x0), height = Math.min(4000, rows.start(order.sourceToDisplay(rect.maxRow) + 1) - y0)
        const canvas = createCanvas(Math.ceil(width), Math.ceil(height))
        const result = renderSheetViewport(sheet, canvas.getContext('2d') as never, geometry, { box: { x: x0, y: y0, width, height }, zoom: 1, deviceScale: 1 }, {}, { cache: new LayoutCache() })
        expect(result.kind, `${name} ${sheet.name}`).toBe('detail')
        // Every saved, non-empty cell of the region is laid out with its text.
        const shown = new Set(result.records.filter(r => r.body.formattedText).map(r => `${r.source.row}:${r.source.col}`))
        const saved = sheet.rows.flatMap(row => row.cells).filter(cell => cell.row >= rect.minRow && cell.row <= rect.maxRow && cell.col >= rect.minCol && cell.col <= rect.maxCol && cell.value !== null && cell.value !== undefined && String(cell.value) !== '')
        for (const cell of saved) expect(shown.has(`${cell.row}:${cell.col}`), `${name} ${sheet.name} ${cell.ref || `${cell.row}:${cell.col}`}`).toBe(true)
        if (saved.length) {
          const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data
          let ink = 0
          for (let i = 0; i < data.length; i += 4) if (data[i]! < 200 || data[i + 1]! < 200 || data[i + 2]! < 200) ink++
          expect(ink, `${name} ${sheet.name}: painted`).toBeGreaterThan(0)
        }
      }
    }
    expect(regions, `${name}: pivots or tables`).toBeGreaterThan(0)
    // The static painter (page mode, export) paints every sheet too.
    const pages = await getPaintables(doc)
    try {
      for (const page of pages) {
        const canvas = createCanvas(Math.max(1, Math.min(4000, Math.ceil(page.spec.widthPx))), Math.max(1, Math.min(4000, Math.ceil(page.spec.heightPx))))
        page.paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
      }
    } finally { pages.dispose() }
  }, 60_000)
})
