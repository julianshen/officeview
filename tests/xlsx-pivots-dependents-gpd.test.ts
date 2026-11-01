// GETPIVOTDATA against the committed layout (Task 31, spec §7.5; Ruling S3-T): measured shapes evaluate, unmeasured
// ones keep their saved values.
import { describe, expect, it, vi } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerGetPivotData, getPivotData, type LivePivotDataContext } from '../src/xlsx/pivots/session/getpivotdata'
import type { PivotLayout } from '../src/xlsx/pivots/engine/place'
import type { XlsxDocument } from '../src/xlsx/types'
import { load, cellAt, addCell, liveContext, viewerReadCellValue, changedOf, F } from './helpers/xlsx-pivots-dependent-fixtures'

registerGetPivotData()

async function run(formulas: Array<[sheet: string, ref: string, col: number, row: number, formula: string]>, opts: { hide?: string; viewerShape?: boolean; file?: string } = {}) {
  const doc = await load(opts.file)
  for (const [sheet, ref, col, row, formula] of formulas) addCell(doc, sheet, ref, col, row, formula, -123456789)
  evaluateWorkbookFormulas(doc, { forceRecalc: true }) // sidestep F1 so only GETPIVOTDATA semantics are probed
  const { ctx } = await liveContext(opts.hide, opts.file)
  const contexts = [ctx]
  const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
    // viewerShape: InteractiveSheetDoc.tsx passes no current sheet (getPivotData(contexts, args, evaluate)).
    getPivotData: (args, evaluate, currentSheet) => getPivotData(contexts, args, evaluate, opts.viewerShape ? currentSheet : 'Formulas'),
    readCellValue: viewerReadCellValue(contexts),
  })
  return { doc, result }
}

describe('GETPIVOTDATA', () => {
  it('looks its cell up by position in the committed layout, never by a scan of the layout per formula', async () => {
    const doc = await load()
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const { ctx } = await liveContext('West')
    const scan = vi.fn(() => undefined)
    // Every lookup that scanned the layout's cells would call this `find`.
    const context = { ...ctx, layout: { ...ctx.layout, cells: Object.assign([...ctx.layout.cells], { find: scan }) } }
    recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
      getPivotData: (args, evaluate) => getPivotData([context], args, evaluate, 'Formulas'),
      readCellValue: viewerReadCellValue([ctx]),
    })
    expect(cellAt(doc, 'Formulas', 'B1')!.value).toBeCloseTo(161104.2, 6)
    expect(cellAt(doc, 'Formulas', 'B3')!.value).toBeCloseTo(9585.23, 6)
    expect(scan).not.toHaveBeenCalled()
  })

  it('a same-sheet, unqualified anchor (what Excel writes) evaluates in the viewer, not #REF!', async () => {
    const { doc } = await run([['Pivot', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",$A$3)']], { hide: 'West', viewerShape: true })
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(161104.2, 6)
  })

  it('a sheet-qualified anchor works through the same viewer call shape', async () => {
    // control: the shipped oracle formulas are qualified (Pivot!$A$3) and pass through the viewer shape.
    const { doc } = await run([['Formulas', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3)']], { hide: 'West', viewerShape: true })
    expect(cellAt(doc, 'Formulas', 'K1')!.value).toBeCloseTo(161104.2, 6)
  })

  it('a formula reading a GETPIVOTDATA cell recomputes with it', async () => {
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'B1*2']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'B1')!.value).toBeCloseTo(161104.2, 6)
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBeCloseTo(322208.4, 6)
  })

  it('S3-T: an unmeasured field on no axis keeps its saved value', async () => {
    const { doc } = await run([
      ['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Channel","NoSuchItem")'],
      ['Formulas', 'C2', 2, 1, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Channel","Web")'],
    ], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBe(-123456789)
    expect(cellAt(doc, 'Formulas', 'C2')!.value).toBe(-123456789)
  })

  it('S3-T: an inner item unique to one parent ("Elderberries") does not get a value the measured rule says is #REF!', async () => {
    // Measured: Product "Cherries" alone is #REF!. "Elderberries" (only under Central) is the same shape: an inner
    // item with no line of its own. The adapter's "single candidate" rule returns 926 instead.
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Product","Elderberries")']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).not.toBe(926)
  })

  it('a Values-only column axis (two data fields, no column field) resolves Region totals', async () => {
    const file = F + 'engine/layout-values-on-columns-no-column-field.xlsx'
    const doc = await load(file)
    addCell(doc, 'Pivot', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East")', -1)
    addCell(doc, 'Pivot', 'K2', 10, 1, 'GETPIVOTDATA("Sum of Units",$A$3)', -1)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const { ctx } = await liveContext(undefined, file)
    recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), { getPivotData: (args, evaluate) => getPivotData([ctx], args, evaluate, 'Pivot') })
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(48118.96, 6)
    expect(cellAt(doc, 'Pivot', 'K2')!.value).toBe(2417)
  })

  it('a data_field given as the source field name ("Sales", Microsoft\'s documented form) is not #REF!', async () => {
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sales",Pivot!$A$3)']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).not.toBe('#REF!')
  })

  it('S3-T: a non-text field-name argument keeps its saved value', async () => {
    const { doc, result } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,42,"East")']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBe(-123456789)
    expect(result.stale.some(key => key.endsWith('!2:0'))).toBe(true)
  })

  it('S3-T: unmeasured GETPIVOTDATA shapes retain their saved values instead of inventing #REF!', async () => {
    const formulas = [
      ['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Unknown data",Pivot!$A$3)'],
      ['Formulas', 'C2', 2, 1, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3:$A$3)'],
      ['Formulas', 'C3', 2, 2, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region")'],
      ['Formulas', 'C4', 2, 3, 'GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Year","2023")'],
    ] as Array<[string, string, number, number, string]>
    const { doc } = await run(formulas, { hide: 'West' })
    for (const [, ref] of formulas) expect(cellAt(doc, 'Formulas', ref)!.value).toBe(-123456789)
  })
})

describe('GETPIVOTDATA on an axis with no fields (its sole data line)', () => {
  const SINGLE = `${F}engine/layout-classic-single-data.xlsx`
  /** The fixture's saved data cell at a row line of its layout (Excel's own saved values). */
  const savedAt = (doc: XlsxDocument, layout: PivotLayout, line: number) => cellAt(doc, 'Pivot', `B${layout.rect.minRow + 2 + line + 1}`)!.value as number
  async function single(formulas: Array<[ref: string, row: number, formula: string]>, change?: (session: Awaited<ReturnType<typeof liveContext>>['session']) => void) {
    const doc = await load(SINGLE)
    for (const [ref, row, formula] of formulas) addCell(doc, 'Pivot', ref, 10, row, formula, -123456789)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const live = await liveContext(undefined, SINGLE)
    change?.(live.session)
    const ctx = { ...live.ctx, def: live.session.def, layout: live.session.layout, rects: [live.savedRect, live.session.layout.rect].map(r => ({ r0: r.minRow, r1: r.maxRow, c0: r.minCol, c1: r.maxCol })) }
    const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
      getPivotData: (args, evaluate, currentSheet, markStale) => getPivotData([ctx], args, evaluate, currentSheet, markStale),
      readCellValue: viewerReadCellValue([ctx]),
    })
    return { doc, ctx, result, saved: live.loaded }
  }
  const formulas: Array<[string, number, string]> = [['K1', 0, 'GETPIVOTDATA("Sum of Sales",$A$3)'], ['K2', 1, 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East")']]

  it('a single-measure pivot with no column fields: the grand total and an item total equal Excel\'s saved cells', async () => {
    const { doc, ctx } = await single(formulas)
    const grand = ctx.layout.axes.rows.findIndex(line => line.type === 'grand'), east = 0
    expect(ctx.layout.axes.cols).toEqual([expect.objectContaining({ type: 'data', items: [] })])
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(savedAt(doc, ctx.layout, grand), 6)
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(222603.36, 6)
    expect(cellAt(doc, 'Pivot', 'K2')!.value).toBeCloseTo(savedAt(doc, ctx.layout, east), 6)
  })

  it('after a change that keeps the totals (Region sorted descending), the values still equal the saved totals', async () => {
    const saved = await single(formulas)
    const east = saved.ctx.layout.axes.rows.findIndex(line => line.items.length === 1 && line.items[0] === 0)
    const savedEast = savedAt(saved.doc, saved.ctx.layout, east)
    const { doc, ctx } = await single(formulas, session => {
      const sorted = session.apply({ kind: 'sort', field: 0, order: 'descending' })
      if (!sorted.ok) throw new Error(sorted.diagnostic.message)
      expect(session.commit(sorted.value)).toBe(true)
    })
    expect(ctx.layout.axes.rows[0]!.items).not.toEqual([0])   // the order changed
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(222603.36, 6)
    expect(cellAt(doc, 'Pivot', 'K2')!.value).toBeCloseTo(savedEast, 6)
  })

  /** The fixture's pivot with Region on columns and no row fields: one data line on rows, a line per Region. */
  function transposed(ctx: LivePivotDataContext & { rects: { r0: number; r1: number; c0: number; c1: number }[] }, doc: XlsxDocument, rows: PivotLayout['axes']['rows'] = [{ type: 'data', items: [], data: 0 } as never]) {
    const regions = ctx.layout.axes.rows.map((line, index) => ({ line, index })).filter(({ line }) => line.items.length === 1 || line.type === 'grand')
    const rect = { minRow: 2, minCol: 0, maxRow: 4, maxCol: regions.length }
    const value = (index: number) => savedAt(doc, ctx.layout, index)
    const layout = { ...ctx.layout, rect, merges: [], axes: { ...ctx.layout.axes, rows, cols: regions.map(({ line }) => line) },
      cells: regions.map(({ index }, k) => ({ row: 4, col: 1 + k, value: { kind: 'number' as const, value: value(index) } })) } as PivotLayout
    const def = { ...ctx.def, rowFields: [], colFields: [0], location: { ...ctx.def.location, ref: 'A3:F5', firstHeaderRow: 1, firstDataRow: 2, firstDataCol: 1 } }
    return { ...ctx, def, layout, layoutCells: undefined, rects: [{ r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol }] }
  }
  const evaluateWith = (doc: XlsxDocument, context: ReturnType<typeof transposed>) => recomputeDependents(doc, changedOf(doc, 'Pivot', context.rects), {
    getPivotData: (args, evaluate, currentSheet, markStale) => getPivotData([context], args, evaluate, currentSheet, markStale),
    readCellValue: () => undefined,
  })

  it('mirrored: no row fields, the sole data line on rows', async () => {
    const { doc, ctx } = await single([])
    const east = ctx.layout.axes.rows.findIndex(line => line.items.length === 1 && line.items[0] === 0)
    const savedEast = savedAt(doc, ctx.layout, east)
    for (const [ref, row, formula] of formulas) addCell(doc, 'Pivot', ref, 10, row, formula, -1)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    evaluateWith(doc, transposed(ctx, doc))
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(222603.36, 6)
    expect(cellAt(doc, 'Pivot', 'K2')!.value).toBeCloseTo(savedEast, 6)
  })

  it('S3-T: an axis with no fields whose lines are not one data line keeps the saved value (no #REF!) and is flagged', async () => {
    const { doc, ctx } = await single([])
    addCell(doc, 'Pivot', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",$A$3)', 777)
    evaluateWorkbookFormulas(doc)
    const result = evaluateWith(doc, transposed(ctx, doc, []))
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBe(777)
    expect(result.stale).toEqual(expect.arrayContaining(['pivot!10:0']))
  })
})

describe('S3-AO: GETPIVOTDATA on a spilled pivot', () => {
  it('keeps its saved value, marked stale (unmeasured, S3-T)', async () => {
    const doc = await load()
    addCell(doc, 'Formulas', 'Z1', 25, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3)', 777)
    evaluateWorkbookFormulas(doc)
    const { ctx } = await liveContext()
    const spilled = { ...ctx, def: { ...ctx.def, spill: { row: 24, col: 0 } } }
    const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
      getPivotData: (args, evaluate, _currentSheet, markStale) => getPivotData([spilled], args, evaluate, 'Formulas', markStale),
      readCellValue: viewerReadCellValue([ctx]),
    })
    expect(cellAt(doc, 'Formulas', 'Z1')!.value).toBe(777)
    expect(result.stale).toEqual(expect.arrayContaining(['formulas!25:0']))
  })
})

describe('GETPIVOTDATA on an axis that carries Values', () => {
  const ON_ROWS = `${F}engine/layout-values-on-rows-two-data.xlsx`, ON_COLS = `${F}engine/layout-classic-two-data.xlsx`
  async function evaluateOn(file: string, formulas: Array<[ref: string, formula: string]>, change?: (session: Awaited<ReturnType<typeof liveContext>>['session']) => void) {
    const doc = await load(file)
    formulas.forEach(([ref, formula], row) => addCell(doc, 'Pivot', ref, 20, row, formula, -123456789))
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const live = await liveContext(undefined, file)
    change?.(live.session)
    const ctx = { ...live.ctx, def: live.session.def, layout: live.session.layout, rects: [live.savedRect, live.session.layout.rect].map(r => ({ r0: r.minRow, r1: r.maxRow, c0: r.minCol, c1: r.maxCol })) }
    const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
      getPivotData: (args, evaluate, currentSheet, markStale) => getPivotData([ctx], args, evaluate, currentSheet, markStale),
      readCellValue: viewerReadCellValue([ctx]),
    })
    return { doc, result, value: (ref: string) => cellAt(doc, 'Pivot', ref)!.value }
  }
  const saved = async (file: string, ref: string) => cellAt(await load(file), 'Pivot', ref)!.value as number
  const sortRegionDescending = (session: Awaited<ReturnType<typeof liveContext>>['session']) => {
    const sorted = session.apply({ kind: 'sort', field: 0, order: 'descending' })
    if (!sorted.ok) throw new Error(sorted.diagnostic.message)
    expect(session.commit(sorted.value)).toBe(true)
  }

  it('Values on rows: a Region alone is its visible subtotal line, a Region and Product the item\'s line, as Excel saved them', async () => {
    const formulas: Array<[string, string]> = [
      ['U1', 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East")'],
      ['U2', 'GETPIVOTDATA("Sum of Units",$A$3,"Region","East")'],
      ['U3', 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East","Product","Apples")'],
      ['U4', 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East","Year",2023)'],
    ]
    for (const change of [undefined, sortRegionDescending]) {
      const { value } = await evaluateOn(ON_ROWS, formulas, change)
      expect(value('U1')).toBeCloseTo(await saved(ON_ROWS, 'E15'), 6)   // "East Sum of Sales", 48118.96
      expect(value('U1')).toBeCloseTo(48118.96, 6)
      expect(value('U2')).toBe(await saved(ON_ROWS, 'E16'))
      expect(value('U3')).toBeCloseTo(await saved(ON_ROWS, 'E7'), 6)
      expect(value('U4')).toBeCloseTo(await saved(ON_ROWS, 'C15'), 6)
    }
  })

  it('Values on columns: a Region\'s grand total and its cell under a Year, as Excel saved them', async () => {
    // Row 6 is East's; H is the Sales grand-total column and E the 2023 Units column (Values inside Year).
    expect(cellAt(await load(ON_COLS), 'Pivot', 'A6')!.value).toBe('East')
    for (const change of [undefined, sortRegionDescending]) {
      const { value } = await evaluateOn(ON_COLS, [
        ['U1', 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East")'],
        ['U2', 'GETPIVOTDATA("Sum of Units",$A$3,"Region","East","Year",2023)'],
      ], change)
      expect(value('U1')).toBeCloseTo(await saved(ON_COLS, 'H6'), 6)
      expect(value('U2')).toBe(await saved(ON_COLS, 'E6'))
    }
  })

  it('S3-T: no unmeasured shape is #REF!: each resolves to its cell, or keeps its saved value and is flagged stale', async () => {
    const shapes: Array<[string, string, string]> = [
      // Values on rows with a partial field list: the inner item alone, and a field on no axis.
      [ON_ROWS, 'U1', 'GETPIVOTDATA("Sum of Sales",$A$3,"Product","Apples")'],
      [ON_ROWS, 'U2', 'GETPIVOTDATA("Sum of Sales",$A$3,"Product","Apples","Year",2023)'],
      [ON_ROWS, 'U3', 'GETPIVOTDATA("Sum of Sales",$A$3,"Channel","Web")'],
      // Two data fields, and a data_field that matches neither.
      [ON_COLS, 'U1', 'GETPIVOTDATA("Sum of Price",$A$3,"Region","East")'],
      [ON_COLS, 'U2', 'GETPIVOTDATA("Average of Sales",$A$3)'],
      // A page field in the arguments.
      [`${F}engine/quirk-page-single.xlsx`, 'U1', 'GETPIVOTDATA("Sum of Sales",$A$3,"Product","Apples")'],
      [`${F}engine/quirk-page-single.xlsx`, 'U2', 'GETPIVOTDATA("Sum of Sales",$A$3,"Product","Bananas")'],
    ]
    for (const file of new Set(shapes.map(([file]) => file))) {
      const own = shapes.filter(([candidate]) => candidate === file)
      const { value, result } = await evaluateOn(file, own.map(([, ref, formula]) => [ref, formula]))
      for (const [, ref, formula] of own) {
        const label = `${formula} on ${file.slice(file.lastIndexOf('/') + 1)}`
        expect(value(ref), label).not.toBe('#REF!')
        if (value(ref) === -123456789) expect(result.stale, label).toContain(`pivot!20:${Number(ref.slice(1)) - 1}`)
        else expect(typeof value(ref), label).toBe('number')
      }
    }
  })
})
