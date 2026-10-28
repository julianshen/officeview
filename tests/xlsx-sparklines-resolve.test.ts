// @vitest-environment node
/**
 * Sparkline data resolution (Task 4). Each test cites the "Measured constants" row of
 * docs/validation/xlsx-sparklines.md it pins. Hand-built sheets cover the rules; the committed Excel reference
 * workbooks (tests/fixtures/xlsx-sparklines) prove the parse -> resolve wiring end to end.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
import type { SparklineArea, XlsxSparkline, XlsxSparklineGroup } from '../src/xlsx/sparklines/model'
import { resolveWorkbookSparklines, type ResolvedSparkline } from '../src/xlsx/sparklines/resolve'

type Raw = number | string | boolean | null | { error: string } | { formula: string; value: number | null; cached: boolean }
const cellOf = (row: number, col: number, raw: Raw): XlsxCell => {
  const base = { ref: `${String.fromCharCode(65 + col)}${row + 1}`, row, col, styleIndex: 0 }
  if (raw !== null && typeof raw === 'object' && 'error' in raw) return { ...base, value: raw.error, valueIsError: true, hasCachedValue: true }
  if (raw !== null && typeof raw === 'object') return { ...base, value: raw.value, formula: raw.formula, hasCachedValue: raw.cached }
  return { ...base, value: raw }
}
/** A sheet from a sparse grid: `cells[row][col]`, `undefined` entries are absent cells. */
function sheetOf(name: string, cells: Array<Array<Raw | undefined>>, extra: Partial<XlsxSheet> = {}): XlsxSheet {
  const rows = cells.map((line, r) => ({ index: r, cells: line.flatMap((raw, c) => raw === undefined ? [] : [cellOf(r, c, raw)]) }))
  return { name, rows, cols: [], merges: [], mergeRanges: [], ...extra }
}
const area = (ref: string, sheet?: string): SparklineArea => {
  const [a, b = a] = ref.split(':')
  const parse = (cell: string) => ({ c: cell.charCodeAt(0) - 65, r: Number(cell.slice(1)) - 1 })
  const from = parse(a), to = parse(b)
  return { ...(sheet ? { sheet } : {}), r0: from.r, c0: from.c, r1: to.r, c1: to.c }
}
const spark = (host: string, data: string | SparklineArea): XlsxSparkline => {
  const h = area(host)
  return { host: { row: h.r0, col: h.c0 }, data: typeof data === 'string' ? area(data) : data, raw: {} }
}
const groupOf = (sparklines: XlsxSparkline[], over: Partial<XlsxSparklineGroup> = {}): XlsxSparklineGroup => ({
  type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'gap', markers: false, high: false, low: false, first: false, last: false, negative: false,
  displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual',
  colors: {}, sparklines, ...over,
})
const resolve = (sheets: XlsxSheet[], cached = false) => { resolveWorkbookSparklines(sheets, { cached }); return sheets }
const at = (sheet: XlsxSheet, host: string): ResolvedSparkline | undefined => { const h = area(host); return sheet.resolvedSparklines?.get(h.r0)?.get(h.c0) }
const withGroups = (sheet: XlsxSheet, ...groups: XlsxSparklineGroup[]): XlsxSheet => ({ ...sheet, sparklineGroups: groups })

describe('resolveWorkbookSparklines', () => {
  it('row and column ranges in reading order', () => {
    const sheet = withGroups(sheetOf('S', [[1, 2, 3, 'x'], [4], [5], [6]]), groupOf([spark('F1', 'A1:C1'), spark('F2', 'A1:A4')]))
    resolve([sheet])
    expect(at(sheet, 'F1')!.values).toEqual([1, 2, 3])
    expect(at(sheet, 'F2')!.values).toEqual([1, 4, 5, 6])
    expect(at(sheet, 'F1')!.group).toBe(sheet.sparklineGroups![0])
    expect(at(sheet, 'F1')).toMatchObject({ firstIndex: 0, lastIndex: 2, min: 1, max: 3 })
  })

  describe('empty cells (gapRule / zeroRule / spanRule)', () => {
    const cells = [[3, 7, undefined, 4, undefined, undefined, 8, 5]]
    const run = (mode: 'gap' | 'zero' | 'span') => { const s = withGroups(sheetOf('S', cells), groupOf([spark('J1', 'A1:H1')], { displayEmptyCellsAs: mode })); resolve([s]); return at(s, 'J1')! }
    it('gap: null slots are kept (slots do not close up), excluded from min/max', () => {
      const r = run('gap')
      expect(r.values).toEqual([3, 7, null, 4, null, null, 8, 5])
      expect(r.bridged).toEqual([])
      expect(r).toMatchObject({ min: 3, max: 8, lowIndices: [0], highIndices: [6] })
    })
    it('zero: an empty cell is the value 0 and takes part in min/max', () => {
      const r = run('zero')
      expect(r.values).toEqual([3, 7, 0, 4, 0, 0, 8, 5])
      expect(r).toMatchObject({ min: 0, max: 8, lowIndices: [2, 4, 5] })
    })
    it('span: null slots are kept and the span flag is carried on the group (the geometry bridges them)', () => {
      const r = run('span')
      expect(r.values).toEqual([3, 7, null, 4, null, null, 8, 5])
      expect(r.group.displayEmptyCellsAs).toBe('span')
      expect(r.bridged).toEqual([])
    })
    it('leading and trailing empties stay null slots (spanRule: not extended)', () => {
      const s = withGroups(sheetOf('S', [[undefined, 3, 7, undefined]]), groupOf([spark('J1', 'A1:D1')], { displayEmptyCellsAs: 'span' }))
      resolve([s])
      expect(at(s, 'J1')!.values).toEqual([null, 3, 7, null])
    })
  })

  describe('text, boolean and error cells (cellValueRule)', () => {
    const run = (x: Raw | undefined) => { const s = withGroups(sheetOf('S', [[3, 7, x, 4, 8, 2, 6, 5]]), groupOf([spark('J1', 'A1:H1')], { displayEmptyCellsAs: 'gap' })); resolve([s]); return at(s, 'J1')! }
    it('text counts as 0 and takes part in min/max', () => {
      const r = run('hello')
      expect(r.values[2]).toBe(0)
      expect(r.min).toBe(0)
      expect(r.lowIndices).toEqual([2])
    })
    it('TRUE counts as 1 and takes part in min/max', () => {
      const r = run(true)
      expect(r.values[2]).toBe(1)
      expect(r.min).toBe(1)
    })
    it('FALSE counts as 0 (not measured; the closest rule, booleans are 1/0)', () => {
      expect(run(false).values[2]).toBe(0)
    })
    it('#N/A is skipped and bridged even with displayEmptyCellsAs="gap", and is excluded from min/max', () => {
      const r = run({ error: '#N/A' })
      expect(r.group.displayEmptyCellsAs).toBe('gap')
      expect(r.values).toEqual([3, 7, null, 4, 8, 2, 6, 5])
      expect(r.bridged).toEqual([2])
      expect(r).toMatchObject({ min: 2, max: 8 })
    })
    it('other error codes (not measured) are treated like #N/A', () => {
      const r = run({ error: '#DIV/0!' })
      expect(r.values[2]).toBeNull()
      expect(r.bridged).toEqual([2])
    })
    it('an empty cell is not bridged under gap, so the two are told apart', () => {
      expect(run(undefined).bridged).toEqual([])
    })
  })

  describe('hidden data', () => {
    const column = [[4], [7], [9], [2], [5], [6]]
    it('hidden rows are removed and the rest re-spaced, not gaps (hiddenRowsRule)', () => {
      const s = withGroups(sheetOf('S', column, { rows: sheetOf('S', column).rows.map(r => ({ ...r, hidden: r.index === 1 || r.index === 3 })) }), groupOf([spark('C1', 'A1:A6')]))
      resolve([s])
      expect(at(s, 'C1')!.values).toEqual([4, 9, 5, 6])
    })
    it('hidden columns are removed and the rest re-spaced (hiddenColsRule)', () => {
      const s = withGroups(sheetOf('S', [[4, 7, 3, 9, 2, 8]], { cols: [{ min: 2, max: 2, hidden: true }, { min: 4, max: 5, hidden: true }] }), groupOf([spark('J1', 'A1:F1')]))
      resolve([s])
      expect(at(s, 'J1')!.values).toEqual([4, 7, 9])
    })
    it('displayHidden draws hidden rows and columns as if visible (displayHiddenRule)', () => {
      const base = sheetOf('S', column)
      const s = withGroups({ ...base, rows: base.rows.map(r => ({ ...r, hidden: r.index === 1 })), cols: [{ min: 0, max: 0, hidden: true }] }, groupOf([spark('C1', 'A1:A6')], { displayHidden: true }))
      resolve([s])
      expect(at(s, 'C1')!.values).toEqual([4, 7, 9, 2, 5, 6])
    })
    it('a stray hidden duplicate row record does not hide data (viewport rule: every record must be hidden)', () => {
      const base = sheetOf('S', column)
      const s = withGroups({ ...base, rows: [...base.rows, { index: 1, hidden: true, cells: [] }] }, groupOf([spark('C1', 'A1:A6')]))
      resolve([s])
      expect(at(s, 'C1')!.values).toEqual([4, 7, 9, 2, 5, 6])
    })
    it('a sparkline whose points are all hidden is not drawn', () => {
      const base = sheetOf('S', [[1], [2]])
      const s = withGroups({ ...base, rows: base.rows.map(r => ({ ...r, hidden: true })) }, groupOf([spark('C1', 'A1:A2')]))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
    })
  })

  describe('reference workbooks (parse -> resolve, both value policies)', () => {
    const path = (id: string) => fileURLToPath(new URL(`./fixtures/xlsx-sparklines/${id}.xlsx`, import.meta.url))
    const open = async (id: string, policy: 'engine' | 'cached') => {
      const doc = await parseXlsx(await OfficePackage.load(readFileSync(path(id))), { valuePolicy: policy })
      return doc
    }
    const first = (doc: Awaited<ReturnType<typeof open>>, row = 0): ResolvedSparkline => doc.sheets[0].resolvedSparklines!.get(row)!.get(9)!
    for (const policy of ['engine', 'cached'] as const) {
      it(`${policy}: filter-hidden rows are skipped exactly like hand-hidden rows (filterHiddenRule, hiddenRowsRule)`, async () => {
        // Excel: the autofilter case draws 5 points [4,9,2,5,6]; the hand-hidden case draws [4,7,9,2,5,6].
        expect(first(await open('line-filter-hidden', policy)).values).toEqual([4, 9, 2, 5, 6])
        expect(first(await open('line-hidden-rows', policy)).values).toEqual([4, 7, 9, 2, 5, 6])
        expect(first(await open('line-hidden-cols', policy)).values).toEqual([4, 7, 9, 2, 5, 6])
      })
      it(`${policy}: text, TRUE and #N/A rows of line-text-bool-error (cellValueRule)`, async () => {
        const doc = await open('line-text-bool-error', policy)
        expect(first(doc, 0).values).toEqual([3, 7, 0, 4, 8, 2, 6, 5])
        expect(first(doc, 1).values).toEqual([3, 7, 1, 4, 8, 2, 6, 5])
        expect(first(doc, 2)).toMatchObject({ values: [3, 7, null, 4, 8, 2, 6, 5], bridged: [2] })
      })
      it(`${policy}: data on a later sheet resolves (laterSheetRule)`, async () => {
        const doc = await open('line-later-sheet', policy)
        const sheet1 = doc.sheets.find(s => s.name === 'Sheet1')!
        expect(doc.sheets.some(s => s.name === 'Sheet3')).toBe(true)
        expect(sheet1.resolvedSparklines!.get(0)!.get(9)!.values).toEqual([4, 7, 3, 9, 2, 8, 5, 6])
        expect(sheet1.diagnostics ?? []).toEqual([])
      })
    }
  })

  describe('data on a later sheet', () => {
    it('resolves, with the sheet name matched case-insensitively', () => {
      const host = withGroups(sheetOf('Sheet1', [[]]), groupOf([spark('A1', area('A1:C1', 'sheet3'))]))
      const data = sheetOf('Sheet3', [[1, 2, 3]])
      resolve([host, sheetOf('Sheet2', []), data])
      expect(at(host, 'A1')!.values).toEqual([1, 2, 3])
      expect(host.diagnostics).toBeUndefined()
    })
    it('reads hidden flags from the data sheet, not the host sheet', () => {
      const host = withGroups(sheetOf('Sheet1', [[]]), groupOf([spark('A1', area('A1:A3', 'Data'))]))
      const data = sheetOf('Data', [[1], [2], [3]])
      data.rows[1].hidden = true
      resolve([host, data])
      expect(at(host, 'A1')!.values).toEqual([1, 3])
    })
  })

  describe('scale', () => {
    it('fails closed for a group-scaled group with a dropped member', () => {
      const s = withGroups(sheetOf('S', [[1, 2], [3, 4]]), groupOf([
        spark('D1', 'A1:B1'), { ...spark('D2', 'A2:B2'), data: undefined },
      ], { minAxisType: 'group', maxAxisType: 'group' }))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
      expect(s.diagnostics?.filter(item => item.feature === 'sparkline-group-incomplete')).toHaveLength(1)
    })

    it('fails closed for custom effective min >= max but preserves flat individual scales', () => {
      const custom = withGroups(sheetOf('S', [[3, 3]]), groupOf([spark('D1', 'A1:B1')], { minAxisType: 'custom', manualMin: 3 }))
      resolve([custom])
      expect(custom.resolvedSparklines).toBeUndefined()
      expect(custom.diagnostics?.filter(item => item.feature === 'sparkline-custom-scale')).toHaveLength(1)
      const individual = withGroups(sheetOf('I', [[3, 3]]), groupOf([spark('D1', 'A1:B1')]))
      resolve([individual])
      expect(at(individual, 'D1')).toMatchObject({ min: 3, max: 3 })
    })

    it('keeps the first duplicate host and rejects non-anchor merged hosts', () => {
      const duplicate = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('D1', 'A1:A1'), spark('D1', 'B1:B1')]))
      resolve([duplicate])
      expect(at(duplicate, 'D1')?.values).toEqual([1])
      expect(duplicate.diagnostics?.filter(item => item.feature === 'sparkline-duplicate-host')).toHaveLength(1)
      const merged = withGroups(sheetOf('M', [[1, 2, 3]], { mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 0, maxCol: 1 }] }), groupOf([spark('B1', 'C1:C1')]))
      resolve([merged])
      expect(merged.resolvedSparklines).toBeUndefined()
      expect(merged.diagnostics?.filter(item => item.feature === 'sparkline-merged-host')).toHaveLength(1)
    })
    it('individual: each sparkline uses its own min and max (individualScaleRule)', () => {
      const s = withGroups(sheetOf('S', [[1, 8], [10, 80]]), groupOf([spark('D1', 'A1:B1'), spark('D2', 'A2:B2')]))
      resolve([s])
      expect(at(s, 'D1')).toMatchObject({ min: 1, max: 8 })
      expect(at(s, 'D2')).toMatchObject({ min: 10, max: 80 })
    })
    it('custom: min and max are independent of each other (customScaleRule); values outside are not dropped (customClipRule)', () => {
      const s = withGroups(sheetOf('S', [[-5, 0, 5, 10]]), groupOf([spark('D1', 'A1:D1')], { minAxisType: 'custom', manualMin: 2, maxAxisType: 'individual' }))
      resolve([s])
      expect(at(s, 'D1')).toMatchObject({ min: 2, max: 10, values: [-5, 0, 5, 10] })
      const t = withGroups(sheetOf('S', [[1, 2, 8]]), groupOf([spark('D1', 'A1:C1')], { maxAxisType: 'custom', manualMax: 12 }))
      resolve([t])
      expect(at(t, 'D1')).toMatchObject({ min: 1, max: 12 })
    })
    // groupScaleRule, measured (docs/validation/xlsx-sparklines.md; line-scale-group): min and max are shared
    // across every sparkline of the group and manual values are ignored under group.
    it('group: min and max over every sparkline in the group, manual values ignored (groupScaleRule, measured)', () => {
      const s = withGroups(sheetOf('S', [[1, 8], [10, 80], [-5, 5]]), groupOf([spark('D1', 'A1:B1'), spark('D2', 'A2:B2'), spark('D3', 'A3:B3')], { minAxisType: 'group', maxAxisType: 'group' }))
      resolve([s])
      for (const host of ['D1', 'D2', 'D3']) expect(at(s, host)).toMatchObject({ min: -5, max: 80 })
      // high/low stay per sparkline.
      expect(at(s, 'D1')).toMatchObject({ highIndices: [1], lowIndices: [0] })
    })
    it('group scale ignores hidden and empty points (Review Focus 2)', () => {
      // Sparkline 1 is the column A1:A5 with a hidden row holding 1000 and an empty cell; sparkline 2 is C1:D1.
      const base = sheetOf('S', [[1, undefined, 7, 2], [1000], [3], [undefined], [-4]])
      base.rows[1].hidden = true
      const s = withGroups(base, groupOf([spark('F1', 'A1:A5'), spark('F2', 'C1:D1')], { minAxisType: 'group', maxAxisType: 'group', displayEmptyCellsAs: 'gap' }))
      resolve([s])
      expect(at(s, 'F1')!.values).toEqual([1, 3, null, -4])
      expect(at(s, 'F2')!.values).toEqual([7, 2])
      for (const host of ['F1', 'F2']) {
        expect(at(s, host)!.max).not.toBe(1000)
        expect(at(s, host)).toMatchObject({ min: -4, max: 7 })
      }
    })
    it('group scale with hidden rows and empty cells together: empties are excluded under gap, counted as 0 under zero', () => {
      const run = (mode: 'gap' | 'zero') => {
        const s = withGroups(sheetOf('S', [[5, 9, undefined, 6], [20, 30, 40, 50]]), groupOf([spark('F1', 'A1:D1'), spark('F2', 'A2:D2')], { minAxisType: 'group', maxAxisType: 'group', displayEmptyCellsAs: mode }))
        s.cols = [{ min: 1, max: 1, hidden: true }]
        resolve([s])
        return at(s, 'F1')!
      }
      expect(run('gap')).toMatchObject({ min: 5, max: 50 })
      expect(run('zero')).toMatchObject({ min: 0, max: 50 })
    })
    it('a sparkline with no plottable point is not drawn and does not move the group scale', () => {
      const s = withGroups(sheetOf('S', [[2, 4], [undefined, undefined]]), groupOf([spark('D1', 'A1:B1'), spark('D2', 'A2:B2')], { minAxisType: 'group', maxAxisType: 'group' }))
      resolve([s])
      expect(at(s, 'D2')).toBeUndefined()
      expect(at(s, 'D1')).toMatchObject({ min: 2, max: 4 })
    })
  })

  describe('points', () => {
    it('all-negative high/low: high is the largest negative value (LibreOffice tdf#152929 trap)', () => {
      const s = withGroups(sheetOf('S', [[-3, -7, -4, -9, -2, -8, -5, -6]]), groupOf([spark('J1', 'A1:H1')]))
      resolve([s])
      expect(at(s, 'J1')).toMatchObject({ highIndices: [4], lowIndices: [3], min: -9, max: -2 })
    })
    it('ties: every tied point is high (low) (tieRule)', () => {
      const s = withGroups(sheetOf('S', [[2, 9, 4, 9, 1, 6, 1, 9]]), groupOf([spark('J1', 'A1:H1')]))
      resolve([s])
      expect(at(s, 'J1')).toMatchObject({ highIndices: [1, 3, 7], lowIndices: [4, 6] })
    })
    it('flat series: resolution yields every point as both high and low (flatSeriesRule is geometry)', () => {
      const s = withGroups(sheetOf('S', [[5, 5, 5]]), groupOf([spark('J1', 'A1:C1')]))
      resolve([s])
      expect(at(s, 'J1')).toMatchObject({ min: 5, max: 5, highIndices: [0, 1, 2], lowIndices: [0, 1, 2] })
    })
    it('a single point resolves (singlePointRule is geometry)', () => {
      const s = withGroups(sheetOf('S', [[5]]), groupOf([spark('J1', 'A1')]))
      resolve([s])
      expect(at(s, 'J1')).toMatchObject({ values: [5], firstIndex: 0, lastIndex: 0, highIndices: [0], lowIndices: [0] })
    })
  })

  describe('same values as the grid', () => {
    it('hand-built formula cells: cached shows the saved value, an uncached formula is blank; engine shows the published value', () => {
      const cells = (): XlsxSheet => withGroups(sheetOf('S', [[
        { formula: '1+1', value: 99, cached: true }, // saved 99; an evaluator would have published 2
        { formula: '2+2', value: null, cached: false },
        7,
      ]]), groupOf([spark('J1', 'A1:C1')]))
      const cached = cells(); resolve([cached], true)
      expect(at(cached, 'J1')!.values).toEqual([99, null, 7])
      const engine = cells(); engine.rows[0].cells[0].value = 2; engine.rows[0].cells[1].value = 4
      resolve([engine], false)
      expect(at(engine, 'J1')!.values).toEqual([2, 4, 7])
    })

    it('parseXlsx: a cached formula whose saved value differs from its recalculated value resolves to what the grid shows', async () => {
      const zip = await JSZip.loadAsync(readFileSync(fileURLToPath(new URL('./fixtures/xlsx-sparklines/line-basic.xlsx', import.meta.url))))
      const sheetPath = 'xl/worksheets/sheet1.xml'
      const xml = (await zip.file(sheetPath)!.async('string'))
        .replace('<c r="A1"><v>4</v></c>', '<c r="A1"><f>1+1</f><v>99</v></c>')
        .replace('<c r="B1"><v>7</v></c>', '<c r="B1"><f>2+2</f></c>')
      zip.file(sheetPath, xml)
      const bytes = await zip.generateAsync({ type: 'uint8array' })
      for (const policy of ['cached', 'engine'] as const) {
        const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: policy })
        const sheet = doc.sheets[0]
        const grid = (ref: string) => sheet.rows[0].cells.find(c => c.ref === ref)!
        const resolved = sheet.resolvedSparklines!.get(0)!.get(9)!
        if (policy === 'cached') {
          // The grid shows the saved 99 for A1 and nothing for the uncached B1.
          expect(grid('A1')).toMatchObject({ value: 99, hasCachedValue: true })
          expect(resolved.values.slice(0, 3)).toEqual([99, null, 3])
        } else {
          // The grid shows whatever the evaluator published; the sparkline reads the same cells.
          expect(resolved.values.slice(0, 2)).toEqual([grid('A1').value, grid('B1').value])
          expect(grid('B1').value).toBe(4)
        }
      }
    })
  })

  describe('staging gate (spec 10)', () => {
    for (const [label, over] of [['dateAxis', { dateAxis: true }], ['rightToLeft', { rightToLeft: true }]] as const) {
      it(`${label} groups resolve to nothing and raise no diagnostic`, () => {
        const s = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('J1', 'A1:C1'), spark('J2', area('A1:B2')), spark('J3', area('A1:B1', 'Nowhere'))], over))
        resolve([s])
        expect(s.resolvedSparklines).toBeUndefined()
        expect(s.diagnostics).toBeUndefined()
      })
    }
    for (const type of ['column', 'stacked'] as const) {
      it(`${type} groups resolve like line groups (Stage 2)`, () => {
        const s = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('J1', 'A1:C1')], { type }))
        resolve([s])
        expect(at(s, 'J1')!.values).toEqual([1, 2, 3])
        expect(at(s, 'J1')!.group.type).toBe(type)
        expect(s.diagnostics).toBeUndefined()
      })
    }
    it('a refused group has no entries and adds no diagnostic of its own (parse already diagnosed it)', () => {
      const s = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('J1', 'A1:C1')], { refused: 'sparkline-type' }))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
      expect(s.diagnostics).toBeUndefined()
    })
    it('a line group beside a gated one still resolves', () => {
      const s = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('J1', 'A1:C1')], { dateAxis: true }), groupOf([spark('J2', 'A1:C1')]))
      resolve([s])
      expect(at(s, 'J1')).toBeUndefined()
      expect(at(s, 'J2')!.values).toEqual([1, 2, 3])
    })
  })

  describe('unresolved data: not drawn and one diagnostic', () => {
    it('a missing sheet', () => {
      const s = withGroups(sheetOf('S', [[1, 2, 3]]), groupOf([spark('J1', area('A1:C1', 'Gone')), spark('J2', area('A2:C2', 'Gone'))]))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
      expect(s.diagnostics).toEqual([{ kind: 'unsupported', feature: 'sparkline-data-reference', message: 'Sparkline data sheet not found: Gone' }])
    })
    it('a 2-D range (deduplicated against the parse diagnostic of the same message)', () => {
      const message = 'Sparkline data reference must be one-dimensional'
      const s = withGroups(sheetOf('S', [[1, 2], [3, 4]], { diagnostics: [{ kind: 'unsupported', feature: 'sparkline-data-reference', message }] }), groupOf([spark('J1', area('A1:B2'))]))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
      expect(s.diagnostics).toHaveLength(1)
      const t = withGroups(sheetOf('S', [[1, 2], [3, 4]]), groupOf([spark('J1', area('A1:B2'))]))
      resolve([t])
      expect(t.diagnostics).toEqual([{ kind: 'unsupported', feature: 'sparkline-data-reference', message }])
    })
    it('a sparkline without a readable host or data adds no diagnostic (parse already reported it)', () => {
      const s = withGroups(sheetOf('S', [[1]]), groupOf([{ host: undefined, data: area('A1'), raw: {} }, { host: { row: 0, col: 9 }, data: undefined, raw: {} }]))
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
      expect(s.diagnostics).toBeUndefined()
    })
    it('re-resolving is idempotent and replaces the cache', () => {
      const s = withGroups(sheetOf('S', [[1, 2]]), groupOf([spark('J1', 'A1:B1'), spark('J2', area('A1:B1', 'Gone'))]))
      resolve([s]); resolve([s])
      expect(s.diagnostics).toHaveLength(1)
      expect(at(s, 'J1')!.values).toEqual([1, 2])
      s.sparklineGroups = []
      resolve([s])
      expect(s.resolvedSparklines).toBeUndefined()
    })
  })
})
