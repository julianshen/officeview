// @vitest-environment node
/**
 * Pivot axes and layout to cells (spec §5.2, §5.7, §6.4; plan Task 17). The oracle is Excel's saved layout:
 * the Stage 1 references (tests/fixtures/xlsx-pivots/pivot-*.xlsx, Excel 16.114 zh-TW), the Excel-made corpus
 * files with rowItems, and the Task 15 manifest cases with `task: 17` once that manifest exists. Synthetic view
 * edits below take their expected lines from the saved cells of the same reference (a blank saved value cell is a
 * combination with no records).
 */
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { savedPivotLayout, savedPivotRoles, type PivotLine } from '../src/xlsx/pivots/savedLayout'
import { pivotButtons } from '../src/xlsx/pivots/buttons'
import type { PivotCacheDefinition, PivotItemType, PivotTableDefinition } from '../src/xlsx/pivots/model'
import { isoToSerial } from '../src/xlsx/pivots/engine/records'
import { loadAll, savedCell, type Loaded } from './helpers/xlsx-pivots-load'
import { initialViewState, type FieldState, type PivotViewState } from '../src/xlsx/pivots/engine/view'
import { computeAxes, lineCarriesValues, NO_LINE, type PivotAxes } from '../src/xlsx/pivots/engine/axes'
import { liveDefinition, placeCells, type PivotCellValue, type PivotLayout, type PivotValueGrid } from '../src/xlsx/pivots/engine/place'
import { checkSavedLayout, sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import type { SheetRect } from '../src/xlsx/tableStyles/model'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const CORPUS = resolve(process.cwd(), 'corpus')
const MANIFEST = resolve(FIXTURES, 'engine/manifest.json')

const cached = new Map<string, Promise<Loaded[]>>()
const load = (path: string) => { let p = cached.get(path); if (!p) cached.set(path, p = loadAll(path)); return p }
const ref = async (name: string) => (await load(resolve(FIXTURES, `${name}.xlsx`)))[0]!
const REFERENCES = readdirSync(FIXTURES).filter(name => /^pivot-.*\.xlsx$/.test(name)).sort()

const axesOf = (p: Loaded, view: PivotViewState = initialViewState(p.def, p.cache)) => computeAxes(p.store, p.cache, p.def, view)
function okAxes(p: Loaded, view?: PivotViewState): PivotAxes {
  const result = axesOf(p, view)
  if (!result.ok) throw new Error(`${p.file} ${p.def.name}: ${result.diagnostic.feature}: ${result.diagnostic.message}`)
  return result.value
}
/** A view with some fields' state replaced. */
function withFields(view: PivotViewState, edits: Record<number, Partial<FieldState>>, rest: Partial<PivotViewState> = {}): PivotViewState {
  return { ...view, ...rest, fields: view.fields.map((field, k) => edits[k] ? { ...field, ...edits[k] } : field) }
}
const L = (type: PivotItemType, first: number, items: number[], data = 0): PivotLine => ({ type, first, items, data })
const G = (data = 0): PivotLine => L('grand', 0, [], data)
const blankGrid: PivotValueGrid = { get: () => ({ kind: 'empty' }) }

/** Every non-value cell of the layout (location and page fields) against the saved sheet: text and numbers equal,
 * an implicit caption only where the saved sheet holds text. Returns the differences. */
function labelDiffs(p: Loaded, layout: PivotLayout): string[] {
  const saved = savedPivotLayout(p.def)!
  const at = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell]))
  const diffs: string[] = []
  const check = (row: number, col: number) => {
    if (row >= saved.dataRow0 && row <= saved.rect.maxRow && col >= saved.dataCol0 && col <= saved.rect.maxCol) return
    const cell = at.get(`${row},${col}`), expected = savedCell(p.sheet, row, col)
    const actual: PivotCellValue = cell?.value ?? { kind: 'empty' }
    if (cell?.caption === 'implicit') { if (expected.kind !== 'text') diffs.push(`R${row + 1}C${col + 1}: implicit caption over ${JSON.stringify(expected)}`); return }
    if (JSON.stringify(actual) !== JSON.stringify(expected)) diffs.push(`R${row + 1}C${col + 1}: ${JSON.stringify(actual)} ≠ saved ${JSON.stringify(expected)}`)
  }
  for (const rect of saved.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) check(row, col)
  for (const cell of layout.cells) if (!saved.rects.some(r => cell.row >= r.minRow && cell.row <= r.maxRow && cell.col >= r.minCol && cell.col <= r.maxCol)) diffs.push(`R${cell.row + 1}C${cell.col + 1}: placed outside the saved rectangles`)
  return diffs
}
const ENGINE_PENDING = 'pivot-engine-pending'

describe('computeAxes on the Stage 1 references', () => {
  it.each(REFERENCES)('%s: lines equal the saved rowItems/colItems', async name => {
    const p = (await load(resolve(FIXTURES, name)))[0]!
    const result = axesOf(p)
    // No Stage 1 reference uses a Task 21 feature: each must compute.
    expect(result.ok, result.ok ? '' : result.diagnostic.message).toBe(true)
    if (!result.ok) return
    const saved = savedPivotLayout(p.def)!
    expect(result.value.rows).toEqual(saved.rows)
    expect(result.value.cols).toEqual(saved.cols)
  })

  it('maps every record to the leaf line of its items', async () => {
    const p = await ref('pivot-layout-compact')
    const axes = okAxes(p)
    const region = p.store.columns[0]!, product = p.store.columns[1]!, year = p.store.columns[2]!
    if (region.kind !== 'shared' || product.kind !== 'shared' || year.kind !== 'shared') throw new Error('shared columns expected')
    const pivotItem = (field: number, x: number) => p.def.fields[field]!.items!.findIndex(item => item.x === x)
    expect(axes.rowKey.length).toBe(p.store.count)
    for (let r = 0; r < p.store.count; r++) {
      expect(axes.rows[axes.rowKey[r]!]!.items).toEqual([pivotItem(0, region.index[r]!), pivotItem(1, product.index[r]!)])
      // Values is the deepest column level: a record's column leaf is the data-field-0 line of its year.
      expect(axes.cols[axes.colKey[r]!]).toEqual(L('data', 0, [pivotItem(2, year.index[r]!), 0]))
    }
  })
})

describe('Task 17 behaviours', () => {
  it('compact collapsed subtotal-top lines equal saved rowItems', async () => {
    const compact = await ref('pivot-layout-compact'), collapsed = await ref('pivot-collapsed-item'), top = await ref('pivot-subtotals-top')
    for (const p of [compact, collapsed, top]) expect(okAxes(p).rows).toEqual(savedPivotLayout(p.def)!.rows)
    // The same records with East collapsed in the live view give the collapsed reference's saved lines.
    const view = initialViewState(compact.def, compact.cache)
    const live = okAxes(compact, withFields(view, { 0: { collapsed: view.fields[0]!.collapsed.map((_, k) => k === 0) } }))
    expect(live.rows).toEqual(savedPivotLayout(collapsed.def)!.rows)
    // Records under the collapsed item map to its line.
    const east = live.rows.findIndex(line => line.type === 'data' && line.items.length === 1 && line.items[0] === 0)
    const region = compact.store.columns[0]!
    if (region.kind !== 'shared') throw new Error('shared column expected')
    for (let r = 0; r < compact.store.count; r++) if (region.index[r] === 3) expect(live.rowKey[r]).toBe(east)
    // Subtotals at the bottom (subtotalTop off) give the subtotals-bottom reference's lines.
    const bottom = await ref('pivot-subtotals-bottom')
    expect(okAxes(compact, withFields(view, { 0: { subtotalTop: false } })).rows).toEqual(savedPivotLayout(bottom.def)!.rows)
  })

  it('compact form with three row fields (corpus pivotcompact, Excel-made)', async ({ skip }) => {
    const path = resolve(CORPUS, 'lo-calc-pivotcompact.xlsx')
    if (!existsSync(path)) skip('corpus absent')
    const p = (await load(path))[0]!
    expect(p.def.rowFields).toHaveLength(3)
    expect(okAxes(p).rows).toEqual(savedPivotLayout(p.def)!.rows)
  })

  it('outline and tabular lines and label columns equal saved', async () => {
    for (const name of ['pivot-layout-outline', 'pivot-layout-tabular', 'pivot-repeat-labels']) {
      const p = await ref(name)
      const axes = okAxes(p), saved = savedPivotLayout(p.def)!
      expect(axes.rows, name).toEqual(saved.rows)
      expect(axes.cols, name).toEqual(saved.cols)
      const layout = placeCells(p.def, initialViewState(p.def, p.cache), p.cache, axes, blankGrid, { date1904: p.date1904 })
      // Each row field's labels sit in its saved label column.
      axes.rows.forEach((line, n) => {
        if (line.type !== 'data') return
        const level = line.items.length - 1
        const cell = layout.cells.find(c => c.row === saved.dataRow0 + n && c.col < saved.dataCol0 && c.col === saved.labelCol[level])
        expect(cell, `${name} line ${n}`).toBeDefined()
      })
    }
    // Switching the compact reference's fields to tabular reproduces the tabular reference's lines.
    const compact = await ref('pivot-layout-compact'), tabular = await ref('pivot-layout-tabular')
    const view = initialViewState(compact.def, compact.cache)
    const flat = { compact: false, outline: false }
    expect(okAxes(compact, withFields(view, { 0: flat, 1: flat, 2: flat })).rows).toEqual(savedPivotLayout(tabular.def)!.rows)
  })

  it('values on rows and on columns place the Values level', async () => {
    const rows = await ref('pivot-values-on-rows'), cols = await ref('pivot-layout-compact')
    const onRows = okAxes(rows)
    expect(onRows.rows).toEqual(savedPivotLayout(rows.def)!.rows)
    expect(onRows.cols).toEqual(savedPivotLayout(rows.def)!.cols)
    // The Values level holds the data-field index as its item, and the line's `data` agrees.
    for (const line of onRows.rows) if (line.type === 'data' && line.items.length === 3) expect(line.items[2]).toBe(line.data)
    // Subtotals with Values below them: one bottom line per data field; one grand line per data field.
    expect(onRows.rows.slice(10, 12)).toEqual([L('default', 0, [0], 0), L('default', 0, [0], 1)])
    expect(onRows.rows.slice(-2)).toEqual([G(0), G(1)])
    expect(lineCarriesValues(initialViewState(rows.def, rows.cache), 'rows', onRows.rows[0]!)).toBe(false)
    const onCols = okAxes(cols)
    expect(onCols.cols).toEqual(savedPivotLayout(cols.def)!.cols)
    expect(onCols.cols.slice(0, 2)).toEqual([L('data', 0, [0, 0], 0), L('data', 1, [0, 1], 1)])
    expect(onCols.cols.slice(-2)).toEqual([G(0), G(1)])
  })

  // From the compact reference's saved 2024 column (F): East has no Bananas, North and West no Apples.
  const ONLY_2024_ROWS = [
    L('data', 0, [0]), L('data', 1, [0, 0]), L('data', 1, [0, 2]),
    L('data', 0, [1]), L('data', 1, [1, 1]), L('data', 1, [1, 2]),
    L('data', 0, [2]), L('data', 1, [2, 0]), L('data', 1, [2, 1]), L('data', 1, [2, 2]),
    L('data', 0, [3]), L('data', 1, [3, 1]), L('data', 1, [3, 2]), G(),
  ]
  it('items with no data are dropped unless showAll', async () => {
    const p = await ref('pivot-layout-compact')
    const view = initialViewState(p.def, p.cache)
    // Year items: 0 = 2022, 1 = 2023, 2 = 2024 (pivot item order), 3 = the default subtotal item.
    const only2024 = withFields(view, { 2: { hidden: [true, true, false, false] } })
    const axes = okAxes(p, only2024)
    expect(axes.rows).toEqual(ONLY_2024_ROWS)
    expect(axes.cols).toEqual([L('data', 0, [2, 0], 0), L('data', 1, [2, 1], 1), G(0), G(1)])
    // showAll on Product brings its items back under every region; hidden years stay hidden under showAll.
    const all = okAxes(p, withFields(only2024, { 1: { showAll: true }, 2: { showAll: true, hidden: [true, true, false, false] } }))
    expect(all.rows).toEqual(savedPivotLayout(p.def)!.rows)
    expect(all.cols).toEqual(axes.cols)
    // Records of hidden years have no column line.
    const year = p.store.columns[2]!
    if (year.kind !== 'shared') throw new Error('shared column expected')
    for (let r = 0; r < p.store.count; r++) if (year.index[r] !== 1) expect([axes.rowKey[r], axes.colKey[r]]).toEqual([NO_LINE, NO_LINE])
  })

  it('hidden items and page selections', async () => {
    const p = await ref('pivot-layout-compact')
    const view = initialViewState(p.def, p.cache)
    // Hide East (pivot item 0): its lines go, its records map to no line, the rest keep their saved lines.
    const noEast = okAxes(p, withFields(view, { 0: { hidden: [true, false, false, false, false] } }))
    expect(noEast.rows).toEqual(savedPivotLayout(p.def)!.rows.slice(4))
    const region = p.store.columns[0]!
    if (region.kind !== 'shared') throw new Error('shared column expected')
    for (let r = 0; r < p.store.count; r++) expect(noEast.rowKey[r] === NO_LINE).toBe(region.index[r] === 3)
    // A single page selection filters records: Year moved to a page with 2024 selected gives the 2024-only rows.
    const page = okAxes(p, { ...view, cols: [-2], pages: [{ field: 2, selected: [2] }] })
    expect(page.rows).toEqual(ONLY_2024_ROWS)
    expect(page.cols).toEqual([L('data', 0, [0], 0), L('data', 0, [1], 1)])
    // The page-field reference: Product on a page, (All) selected; selecting Bananas keeps exactly its records.
    const paged = await ref('pivot-page-field')
    const pagedView = initialViewState(paged.def, paged.cache)
    expect(pagedView.pages).toEqual([{ field: 1, selected: 'all' }])
    const bananas = okAxes(paged, { ...pagedView, pages: [{ field: 1, selected: [1] }] })
    const product = paged.store.columns[1]!
    if (product.kind !== 'shared') throw new Error('shared column expected')
    for (let r = 0; r < paged.store.count; r++) expect(bananas.rowKey[r] !== NO_LINE).toBe(product.index[r] === 1)
  })

  it('cross-field visibility: a page selection leaving an item without records drops it (corpus POI sample, Excel-made)', async ({ skip }) => {
    const path = resolve(CORPUS, 'poi-ExcelPivotTableSample.xlsx')
    if (!existsSync(path)) skip('corpus absent')
    // Tabelle3: Field3 on a page with item 1 hidden; Field1's item 1 then has no records and Excel drops it.
    const p = (await load(path)).find(candidate => candidate.def.sheet === 'Tabelle3')!
    expect(p.def.fields[2]!.items!.map(item => item.hidden)).toEqual([false, true, false, false])
    expect(okAxes(p).rows).toEqual([L('data', 0, [0]), L('data', 0, [2]), G()])
    expect(okAxes(p).rows).toEqual(savedPivotLayout(p.def)!.rows)
  })

  it('liveDefinition of the saved state reproduces the saved definition', async () => {
    for (const name of REFERENCES) {
      const p = (await load(resolve(FIXTURES, name)))[0]!
      const view = initialViewState(p.def, p.cache), axes = okAxes(p, view)
      const layout = placeCells(p.def, view, p.cache, axes, blankGrid, { date1904: p.date1904 })
      const live = liveDefinition(p.def, view, axes, layout)
      expect(live.rowItems, name).toEqual(p.def.rowItems)
      expect(live.colItems, name).toEqual(p.def.colItems)
      expect(live.location, name).toEqual(p.def.location)
      const want = savedPivotRoles(p.def)!, got = savedPivotRoles(live)!
      expect(got.rects, name).toEqual(want.rects)
      for (const rect of want.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) {
        expect(got.rowClass(row), `${name} row ${row}`).toBe(want.rowClass(row))
        for (let col = rect.minCol; col <= rect.maxCol; col++) expect(got.roleAt(row, col), `${name} R${row + 1}C${col + 1}`).toEqual(want.roleAt(row, col))
      }
      for (const mode of ['screen', 'print'] as const) expect(pivotButtons(live, mode), `${name} ${mode}`).toEqual(pivotButtons(p.def, mode))
    }
  })

  it('liveDefinition of a changed view moves the buttons and roles with the lines', async () => {
    const p = await ref('pivot-layout-compact'), collapsed = await ref('pivot-collapsed-item')
    const view = initialViewState(p.def, p.cache)
    const changed = withFields(view, { 0: { collapsed: view.fields[0]!.collapsed.map((_, k) => k === 0) } })
    const axes = okAxes(p, changed)
    const live = liveDefinition(p.def, changed, axes, placeCells(p.def, changed, p.cache, axes, blankGrid, { date1904: p.date1904 }))
    expect(live.location).toEqual(collapsed.def.location)
    expect(live.rowItems).toEqual(collapsed.def.rowItems)
    expect(live.fields[0]!.items![0]!.showDetails).toBe(false)
    for (const mode of ['screen', 'print'] as const) expect(pivotButtons(live, mode)).toEqual(pivotButtons(collapsed.def, mode))
  })

  it('named subtotals add one line per function', async () => {
    // Unmeasured until Task 15's named-subtotal case: one line per function in schema order, at the bottom even
    // with subtotalTop (a parent line can hold one value per cell), each with the function as its type.
    const p = await ref('pivot-layout-compact')
    const view = withFields(initialViewState(p.def, p.cache), { 0: { defaultSubtotal: false, subtotals: ['sum', 'count', 'avg'] } })
    const rows = okAxes(p, view).rows
    const expected: PivotLine[] = []
    for (let region = 0; region < 4; region++) {
      expected.push(L('data', 0, [region]), L('data', 1, [region, 0]), L('data', 1, [region, 1]), L('data', 1, [region, 2]))
      expected.push(L('sum', 0, [region]), L('count', 0, [region]), L('avg', 0, [region]))
    }
    expect(rows).toEqual([...expected, G()])
    // The parent line then carries no values; each function line does.
    expect(lineCarriesValues(view, 'rows', rows[0]!)).toBe(false)
    expect(lineCarriesValues(view, 'rows', rows[4]!)).toBe(true)
  })

  it('placeCells puts labels, captions and blank rows where the saved cells are', async () => {
    for (const name of REFERENCES) {
      const p = (await load(resolve(FIXTURES, name)))[0]!
      const view = initialViewState(p.def, p.cache), axes = okAxes(p, view)
      const layout = placeCells(p.def, view, p.cache, axes, blankGrid, { date1904: p.date1904 })
      const saved = savedPivotLayout(p.def)!
      expect(layout.rect, name).toEqual(saved.rect)
      expect(labelDiffs(p, layout), name).toEqual([])
      // These references are zh-TW: their default subtotal and total captions are implicit.
      expect(layout.captionsUnverified, name).toBe(layout.cells.some(cell => cell.caption === 'implicit'))
      // With the saved caption slots supplied, every caption is the saved text and none is unverified.
      const withSaved = placeCells(p.def, view, p.cache, axes, blankGrid, {
        date1904: p.date1904,
        savedCaption: (row, col) => { const value = savedCell(p.sheet, row, col); return value.kind === 'text' ? value.value : undefined },
      })
      expect(withSaved.captionsUnverified, name).toBe(false)
      expect(labelDiffs(p, withSaved), name).toEqual([])
    }
  })

  it('placeCells asks the value grid for every value cell the saved sheet fills', async () => {
    const asked: PivotValueGrid = { get: (row, col, data) => ({ kind: 'text', value: `${row}/${col}/${data}` }) }
    for (const name of REFERENCES) {
      const p = (await load(resolve(FIXTURES, name)))[0]!
      const view = initialViewState(p.def, p.cache), axes = okAxes(p, view)
      const layout = placeCells(p.def, view, p.cache, axes, asked, { date1904: p.date1904 })
      const saved = savedPivotLayout(p.def)!
      const at = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell]))
      const valuesOnRows = view.rows.includes(-2)
      for (let row = saved.dataRow0; row <= saved.rect.maxRow; row++) for (let col = saved.dataCol0; col <= saved.rect.maxCol; col++) {
        const n = row - saved.dataRow0, m = col - saved.dataCol0
        const cell = at.get(`${row},${col}`)
        const carries = lineCarriesValues(view, 'rows', axes.rows[n]!) && lineCarriesValues(view, 'cols', axes.cols[m]!)
        if (savedCell(p.sheet, row, col).kind !== 'empty') expect(carries, `${name} R${row + 1}C${col + 1}`).toBe(true)
        const data = valuesOnRows ? axes.rows[n]!.data : axes.cols[m]!.data
        expect(cell?.value, `${name} R${row + 1}C${col + 1}`).toEqual(carries ? { kind: 'text', value: `${n}/${m}/${data}` } : undefined)
      }
    }
  })

  it('blank rows: insertBlankRow adds a blank line after each item block, with no cells', async () => {
    // Unmeasured until Task 15's blank-row case: after the item's subtotal, before the next item and the grand total.
    const p = await ref('pivot-layout-compact')
    const view = withFields(initialViewState(p.def, p.cache), { 0: { insertBlankRow: true } })
    const axes = okAxes(p, view)
    const saved = savedPivotLayout(p.def)!.rows
    const expected: PivotLine[] = []
    for (let region = 0; region < 4; region++) expected.push(...saved.slice(4 * region, 4 * region + 4), L('blank', 0, [region]))
    expect(axes.rows).toEqual([...expected, G()])
    const layout = placeCells(p.def, view, p.cache, axes, { get: () => ({ kind: 'number', value: 1 }) }, { date1904: p.date1904 })
    const blankRow = layout.rect.minRow + p.def.location.firstDataRow + 4
    expect(layout.cells.filter(cell => cell.row === blankRow)).toEqual([])
    expect(layout.rect.maxRow).toBe(savedPivotLayout(p.def)!.rect.maxRow + 4)
  })
})

describe('unmeasured layout paths', () => {
  const unmeasured = (result: ReturnType<typeof computeAxes>) => result.ok ? undefined : { kind: result.diagnostic.kind, feature: result.diagnostic.feature }
  const UNMEASURED = { kind: 'unsupported', feature: 'pivot-engine-unmeasured' }

  it('a hidden item on a field on no axis or page fails closed until measured (Task 15)', async () => {
    // Sales (field 3) is on no axis; hiding anything there would filter records by a rule Excel has not shown us.
    // (quirk-off-axis-hidden-item: Excel itself drops a field's hidden items when it leaves every axis.)
    const p = await ref('pivot-layout-compact')
    const view = initialViewState(p.def, p.cache)
    expect(unmeasured(axesOf(p, withFields(view, { 3: { hidden: [true] } })))).toEqual(UNMEASURED)
    // Hidden items on row, column and page fields still filter (covered above): the same edit on Year computes.
    expect(axesOf(p, withFields(view, { 2: { hidden: [true, false, false, false] } })).ok).toBe(true)
  })
})

/** Fix round B1: rules read off the Task 15 Excel cases (tests/fixtures/xlsx-pivots/engine/*.xlsx). */
describe('fix round B1: layout rules measured on the Excel oracle', () => {
  const engine = async (name: string) => (await load(resolve(FIXTURES, 'engine', `${name}.xlsx`)))[0]!
  const refusal = (result: ReturnType<typeof computeAxes>) => result.ok ? undefined : `${result.diagnostic.feature}: ${result.diagnostic.message}`
  const layoutOf = (p: Loaded, view: PivotViewState = initialViewState(p.def, p.cache)) => placeCells(p.def, view, p.cache, okAxes(p, view), blankGrid, { date1904: p.date1904 })
  const at = (layout: PivotLayout, row: number, col: number) => layout.cells.find(cell => cell.row === row && cell.col === col)

  it('hideValuesRow hides the Values row in outline and tabular form, not only compact', async () => {
    // layout-outline-three-rows / layout-tabular-three-rows: Values alone on columns, compactData off, hideValuesRow:
    // Excel saves A3:E60 with firstHeaderRow 0 and firstDataRow 1 (57 lines under one header row).
    for (const name of ['layout-outline-three-rows', 'layout-tabular-three-rows']) {
      const p = await engine(name)
      expect(p.def.compactData, name).toBe(false)
      const view = initialViewState(p.def, p.cache), axes = okAxes(p, view)
      const layout = placeCells(p.def, view, p.cache, axes, blankGrid, { date1904: p.date1904 })
      expect(liveDefinition(p.def, view, axes, layout).location, name).toEqual({ ref: 'A3:E60', firstHeaderRow: 0, firstDataRow: 1, firstDataCol: 3, rowPageCount: 0, colPageCount: 0 })
      expect(labelDiffs(p, layout), name).toEqual([])
    }
  })

  it('classic layout with no column field and one data field: data name top-left, a Total header over the values', async () => {
    // layout-classic-single-data: A3 "Sum of Sales", A4 "Row Labels", B4 "合計" (Excel's Total), data from row 5;
    // location A3:B23, firstHeaderRow 2, firstDataRow 2, firstDataCol 1.
    const p = await engine('layout-classic-single-data')
    expect(p.def.gridDropZones).toBe(true)
    const layout = layoutOf(p)
    expect(layout.rect).toEqual({ minRow: 2, minCol: 0, maxRow: 22, maxCol: 1 })
    expect(at(layout, 2, 0)?.value).toEqual({ kind: 'text', value: 'Sum of Sales' })
    expect(at(layout, 3, 0)?.value).toEqual({ kind: 'text', value: 'Row Labels' })
    expect(at(layout, 3, 1)).toEqual({ row: 3, col: 1, value: { kind: 'text', value: 'Total' }, caption: 'implicit' })
    // Classic without a column field stays refused where no case measures it: two data fields (Values on rows).
    const view = initialViewState(p.def, p.cache)
    const two = { ...view, rows: [...view.rows, -2], data: [view.data[0]!, { ...view.data[0]!, field: 4, name: '' }] }
    expect(refusal(axesOf(p, two))).toBe('pivot-engine-unmeasured: classic pivot layout (gridDropZones) without a column field is measured for one data field with row fields only')
    expect(refusal(axesOf(p, { ...view, rows: [] }))).toBe('pivot-engine-unmeasured: classic pivot layout (gridDropZones) without a column field is measured for one data field with row fields only')
  })

  it('a collapsed item above the Values field: its own line, one line per data field, no subtotal', async () => {
    // layout-collapsed-item-above-values (rows Region, Product, Values; East collapsed): <i><x/></i>, then
    // <i r="2"><x/></i> and <i r="2" i="1"><x v="1"/></i>: the Product level is left unfilled, r is the Values level,
    // and East has no "East Sum of Sales" subtotal lines (sheet rows 5–7, then North).
    const p = await engine('layout-collapsed-item-above-values')
    const axes = okAxes(p)
    expect(axes.rows.slice(0, 4)).toEqual([L('data', 0, [0]), L('data', 2, [0, -1, 0], 0), L('data', 2, [0, -1, 1], 1), L('data', 0, [1])])
    // East's records map to its first data-field line.
    const region = p.store.columns[0]!
    if (region.kind !== 'shared') throw new Error('shared column expected')
    const eastX = p.def.fields[0]!.items![0]!.x
    expect(p.cache.fields[0]!.sharedItems!.items[eastX!]).toEqual({ type: 'string', value: 'East' })
    let eastRecords = 0
    for (let r = 0; r < p.store.count; r++) if (region.index[r] === eastX) { eastRecords++; expect(axes.rowKey[r]).toBe(1) }
    expect(eastRecords).toBeGreaterThan(0)
    // The same rule on the Stage 1 values-on-rows reference (same fields): collapsing its first region.
    const rows = await ref('pivot-values-on-rows')
    const view = initialViewState(rows.def, rows.cache)
    const collapsed = withFields(view, { 0: { collapsed: view.fields[0]!.collapsed.map((_, k) => k === 0) } })
    const live = okAxes(rows, collapsed), saved = savedPivotLayout(rows.def)!.rows
    const next = saved.findIndex(line => line.type === 'data' && line.items.length === 1 && line.items[0] !== 0)
    expect(live.rows).toEqual([L('data', 0, [0]), L('data', 2, [0, -1, 0], 0), L('data', 2, [0, -1, 1], 1), ...saved.slice(next)])
    // Unmeasured variants stay refused: a tabular collapsed field, Values not the last row level, columns.
    const tabular = withFields(collapsed, { 0: { outline: false, compact: false } })
    expect(refusal(axesOf(rows, tabular))).toBe('pivot-engine-unmeasured: a collapsed item above the Values field is measured in outline form only')
    expect(refusal(axesOf(rows, { ...collapsed, rows: [0, -2, 1] }))).toBe('pivot-engine-unmeasured: a collapsed item above the Values field is measured with Values as the last row level only')
    expect(refusal(axesOf(rows, { ...collapsed, rows: [2], cols: [0, 1, -2] }))).toBe('pivot-engine-unmeasured: a collapsed item above the Values field on columns is not measured')
  })

  it('mergeItem: each outer label merged over its block, total captions across the label columns', async () => {
    // layout-merge-labels (tabular Region, Product; columns Year, Values): A6:A8 East over its three products, A9:B9
    // "East 合計", A24:B24 the grand total, C4:D4 2022 over its two data fields, I4:I5 the grand-total column caption.
    const p = await engine('layout-merge-labels')
    const merges = [...layoutOf(p).merges].sort(byRect)
    expect(merges).toEqual(savedMerges(p))
    expect(merges).toContainEqual({ minRow: 5, minCol: 0, maxRow: 7, maxCol: 0 })
    expect(merges).toContainEqual({ minRow: 8, minCol: 0, maxRow: 8, maxCol: 1 })
    expect(merges).toContainEqual({ minRow: 3, minCol: 2, maxRow: 3, maxCol: 3 })
    expect(merges).toContainEqual({ minRow: 3, minCol: 8, maxRow: 4, maxCol: 8 })
    // Central has one product: no single-cell merge.
    expect(merges).not.toContainEqual({ minRow: 21, minCol: 0, maxRow: 21, maxCol: 0 })
    // The Stage 1 tabular reference (same shape) with mergeItem computes; without it there are no merges.
    const tabular = await ref('pivot-layout-tabular')
    expect(layoutOf(tabular).merges).toEqual([])
    const merged = { ...tabular, def: { ...tabular.def, mergeItem: true } }
    expect(layoutOf(merged).merges.length).toBe(4 + 4 + 1 + 3 + 2)
    // Unmeasured shapes stay refused: compact/outline fields (own parent lines).
    const compact = await ref('pivot-layout-compact')
    const compactMerged = { ...compact.def, mergeItem: true }
    expect(refusal(computeAxes(compact.store, compact.cache, compactMerged, initialViewState(compactMerged, compact.cache))))
      .toBe('pivot-engine-unmeasured: merged item labels (mergeItem) are measured for tabular axis fields with totals only after outer items')
  })

  it('an item shown by showAll with no records lists the visible items of the next field', async () => {
    // quirk-cross-hidden-item-showall: Region showAll, Product's Elderberries hidden. Central's only records are
    // Elderberries, yet Excel keeps Central (showAll) and lists Apples, Bananas, Cherries under it (rowItems 16–19).
    const p = await engine('quirk-cross-hidden-item-showall')
    const axes = okAxes(p)
    expect(axes.rows.slice(16)).toEqual([L('data', 0, [4]), L('data', 1, [4, 0]), L('data', 1, [4, 1]), L('data', 1, [4, 2]), G()])
    expect(axes.rows).toEqual(savedPivotLayout(p.def)!.rows)
    // Without showAll Central goes (quirk-cross-hidden-item).
    const plain = await engine('quirk-cross-hidden-item')
    expect(okAxes(plain).rows.some(line => line.items[0] === 4)).toBe(false)
    // Whether such a parent lists an item no passing record has is not measured: hiding East..West leaves
    // Apples..Cherries without records anywhere, so the case fails closed.
    const view = initialViewState(p.def, p.cache)
    const none = withFields(view, { 0: { hidden: view.fields[0]!.hidden.map((_, k) => k < 4) } })
    expect(refusal(axesOf(p, none))).toBe('pivot-engine-unmeasured: an item without records under showAll lists field 1 item 0, which no passing record has')
  })
})

describe('fix round 1', () => {
  it('date items are serials in the workbook date system (1904 as well as 1900)', async () => {
    // Year's shared items replaced by dates: the item labels must follow the date system the caller states.
    const p = await ref('pivot-layout-compact')
    const dates = ['2022-03-01T00:00:00', '2024-03-01T00:00:00', '2023-03-01T00:00:00']
    const cache: PivotCacheDefinition = { ...p.cache, fields: p.cache.fields.map((field, k) => k === 2
      ? { ...field, sharedItems: { ...field.sharedItems!, items: dates.map(value => ({ type: 'date' as const, value })) } } : field) }
    const view = initialViewState(p.def, cache), axes = okAxes({ ...p, cache }, view)
    const label = (date1904: boolean) => placeCells(p.def, view, cache, axes, blankGrid, { date1904 }).cells.find(cell => cell.row === 3 && cell.col === 1)?.value
    // B4 holds the first year item (x=0, 2022-03-01).
    expect(label(false)).toEqual({ kind: 'number', value: isoToSerial(dates[0]!, false) })
    expect(label(true)).toEqual({ kind: 'number', value: isoToSerial(dates[0]!, true) })
    expect(isoToSerial(dates[0]!, false) - isoToSerial(dates[0]!, true)).toBe(1462)
  })

  it('liveDefinition materializes the items of a field saved without an items list when the view edits it', async () => {
    const p = await ref('pivot-layout-compact')
    // Region saved with no items list (other writers): its pivot items are the cache's shared-item order.
    const def: PivotTableDefinition = { ...p.def, fields: p.def.fields.map((field, k) => k === 0 ? (({ items: _items, ...rest }) => rest)(field) : field) }
    const saved = initialViewState(def, p.cache)
    expect(saved.fields[0]!.hidden).toEqual([false, false, false, false])
    // Hide North (cache item 1) and collapse West (cache item 2).
    const edited = withFields(saved, { 0: { hidden: [false, true, false, false], collapsed: [false, false, true, false] } })
    const axes = computeAxes(p.store, p.cache, def, edited)
    if (!axes.ok) throw new Error(axes.diagnostic.message)
    const live = liveDefinition(def, edited, axes.value, placeCells(def, edited, p.cache, axes.value, blankGrid, { date1904: false }))
    expect(live.fields[0]!.items?.map(item => [item.type, item.x, item.hidden, item.showDetails])).toEqual([
      ['data', 0, false, true], ['data', 1, true, true], ['data', 2, false, false], ['data', 3, false, true], ['default', undefined, false, true],
    ])
    // The live definition read back gives the same view and the same axes.
    const reread = initialViewState(live, p.cache)
    expect(reread.fields[0]).toEqual(edited.fields[0] && { ...edited.fields[0], hidden: [...edited.fields[0].hidden, false], collapsed: [...edited.fields[0].collapsed, false] })
    const again = computeAxes(p.store, p.cache, live, reread)
    if (!again.ok) throw new Error(again.diagnostic.message)
    expect(again.value.rows).toEqual(axes.value.rows)
    expect(again.value.cols).toEqual(axes.value.cols)
    // An unedited field without items stays without items.
    const untouched = liveDefinition(def, saved, okAxes({ ...p, def }, saved), placeCells(def, saved, p.cache, okAxes({ ...p, def }, saved), blankGrid, { date1904: false }))
    expect(untouched.fields[0]!.items).toBeUndefined()
  })
})

describe('Task 21 behaviours and remaining gates', () => {
  it('implements sorting and filters, and keeps unmeasured scope/grouping typed', async () => {
    const p = await ref('pivot-layout-compact')
    const view = initialViewState(p.def, p.cache)
    const pending = (result: ReturnType<typeof computeAxes>) => result.ok ? undefined : { kind: result.diagnostic.kind, feature: result.diagnostic.feature }
    const expected = { kind: 'unsupported', feature: 'pivot-engine-unmeasured' }
    expect(axesOf(p, withFields(view, { 0: { sortType: 'descending' } })).ok).toBe(true)
    expect(pending(axesOf(p, withFields(view, { 0: { autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [] } } })))).toEqual(expected)
    expect(axesOf(p, { ...view, filters: [{ fld: 0, type: 'captionEqual', evalOrder: 0, id: 1, stringValue1: 'East', autoFilter: { columns: [] } }] }).ok).toBe(true)
    const grouped: PivotCacheDefinition = { ...p.cache, fields: p.cache.fields.map((field, k) => k === 2 ? { ...field, fieldGroup: { base: 2, rangePr: { autoStart: true, autoEnd: true, groupBy: 'range', groupInterval: 1 } } } : field) }
    const result = computeAxes(p.store, grouped, p.def, view)
    expect(pending(result)).toEqual({ kind: 'parse', feature: 'pivot-definition-invalid' })
    // Calculated fields (Task 22) are evaluated as data only; on an axis or page they are unsupported.
    const calculated: PivotCacheDefinition = { ...p.cache, fields: p.cache.fields.map((field, k) => k === 4 ? { ...field, databaseField: false, formula: 'Sales*2', sharedItems: undefined } : field) }
    expect(computeAxes(p.store, calculated, p.def, view).ok).toBe(true)
    expect(pending(computeAxes(p.store, calculated, p.def, { ...view, cols: [-2, 4] }))).toEqual({ kind: 'unsupported', feature: 'pivot-calculated-field' })
    const formulaOnly: PivotCacheDefinition = { ...p.cache, fields: p.cache.fields.map((field, k) => k === 1 ? { ...field, formula: 'Sales' } : field) }
    expect(pending(computeAxes(p.store, formulaOnly, p.def, view))).toEqual({ kind: 'unsupported', feature: 'pivot-calculated-field' })
    const noFormula: PivotCacheDefinition = { ...p.cache, fields: p.cache.fields.map((field, k) => k === 4 ? { ...field, databaseField: false, sharedItems: undefined } : field) }
    expect(pending(computeAxes(p.store, noFormula, p.def, view))).toEqual({ kind: 'unsupported', feature: ENGINE_PENDING })
    // A sort on a field off every axis changes nothing.
    expect(axesOf(p, withFields(view, { 3: { sortType: 'ascending' } })).ok).toBe(true)
  })
})

/** Corpus files with rowItems: each status as the engine finds it. These are writer files: a disagreement is
 * recorded here and in the task report, never tuned toward. */
const CORPUS_STATUS: Record<string, string[]> = {
  'poi-ExcelPivotTableSample.xlsx': ['equal', 'equal'],
  // Calculated data fields (Task 22). Collabora-written: its Values-only column items are seven `<i/>` with i=0,
  // where Excel writes x=k, i=k, so the computed lines disagree (the writer's own output, not tuned toward).
  'lo-calc-calcfields.xlsx': ['disagrees'],
  'lo-calc-pivot_dark1.xlsx': ['equal'],
  'lo-calc-Pivot2_Row_Compact.xlsx': ['equal'],
  'lo-calc-Pivot2_Row_Grand_Subtotals.xlsx': ['equal'],
  'lo-calc-Pivot3_Column_Grand_Subtotals.xlsx': ['equal'],
  'lo-calc-Pivot4_Column_Grand_Subtotals_SortDescending.xlsx': ['equal'],
  'lo-calc-pivotcompact.xlsx': ['equal'],
  'lo-calc-PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx': ['equal'],
  // Classic layout, no column field, one data field: refused until fix round B1 measured that geometry
  // (layout-classic-single-data); its location (firstHeaderRow 2) and labels now agree too.
  'lo-calc-pivottable_date_field_filter.xlsx': ['equal'],
  'lo-calc-pivottable_outline_mode.xlsx': ['equal'],
  // Field 6, on no axis, has a hidden item: unmeasured, so it fails closed. (Its saved layout shows the record that
  // item would exclude; the file's naming, "Sum - Amount" and "Data", is LibreOffice's.)
  'lo-calc-pivottable_rowcolpage_field_filter.xlsx': ['unmeasured'],
  'lo-calc-pivottable_tabular_mode.xlsx': ['equal'],
  'lo-calc-PivotTableCellFormatsTest_7_TwoRowTwoColumnFields_DataColor.xlsx': ['equal'],
  'lo-calc-PivotTableFormatsSubTotals.xlsx': ['equal'],
  'lo-calc-PivotTableWithNoSourceData.xlsx': ['unmeasured'],
  'lo-calc-shared-dategroup.xlsx': ['equal', 'equal'],
  'lo-calc-shared-numgroup.xlsx': ['equal', 'equal'],
  'lo-calc-two-data-fields.xlsx': ['equal'],
}
describe.skipIf(!existsSync(resolve(CORPUS, 'poi-ExcelPivotTableSample.xlsx')))('corpus files with rowItems', () => {
  it.each(Object.keys(CORPUS_STATUS))('%s', async file => {
    const statuses: string[] = []
    for (const p of await load(resolve(CORPUS, file))) {
      const result = axesOf(p)
      if (!result.ok) {
        const status = ({ [ENGINE_PENDING]: 'pending', 'pivot-engine-unmeasured': 'unmeasured' } as Record<string, string>)[result.diagnostic.feature]
        statuses.push(status ?? result.diagnostic.feature)
        continue
      }
      const saved = savedPivotLayout(p.def)!
      const equal = JSON.stringify([result.value.rows, result.value.cols]) === JSON.stringify([saved.rows, saved.cols])
      statuses.push(equal ? 'equal' : 'disagrees')
      if (!equal) continue
      // Where the lines agree, the labels, captions and location agree too.
      const view = initialViewState(p.def, p.cache)
      const layout = placeCells(p.def, view, p.cache, result.value, blankGrid, { date1904: p.date1904 })
      expect(labelDiffs(p, layout), `${file} ${p.def.name}`).toEqual([])
      expect(liveDefinition(p.def, view, result.value, layout).location, `${file} ${p.def.name}`).toEqual(p.def.location)
    }
    expect(statuses).toEqual(CORPUS_STATUS[file])
  })
})

interface ManifestCase { case: string; file: string; features: string[]; task: number; pivot: string; status?: string }
const manifest: ManifestCase[] | undefined = existsSync(MANIFEST)
  ? (raw => Array.isArray(raw) ? raw : raw.cases)(JSON.parse(readFileSync(MANIFEST, 'utf8')))
  : undefined
const task17 = (manifest ?? []).filter(entry => entry.task === 17 && entry.status !== 'pending')
async function manifestPivot(entry: ManifestCase): Promise<Loaded> {
  const local = resolve(FIXTURES, 'engine', entry.file)
  const all = await load(existsSync(local) ? local : resolve(process.cwd(), entry.file))
  const p = all.find(candidate => candidate.def.name === entry.pivot)
  if (!p) throw new Error(`${entry.case}: no pivot named ${entry.pivot} in ${entry.file}`)
  return p
}
const byRect = (a: SheetRect, b: SheetRect) => a.minRow - b.minRow || a.minCol - b.minCol || a.maxRow - b.maxRow || a.maxCol - b.maxCol
/** The saved sheet's merged ranges inside the pivot's location and page-field cells, sorted. */
function savedMerges(p: Loaded): SheetRect[] {
  const rects = savedPivotLayout(p.def)!.rects
  return (p.sheet.mergeRanges ?? []).filter(m => rects.some(r => m.minRow >= r.minRow && m.maxRow <= r.maxRow && m.minCol >= r.minCol && m.maxCol <= r.maxCol))
    .map(({ minRow, minCol, maxRow, maxCol }) => ({ minRow, minCol, maxRow, maxCol })).sort(byRect)
}
async function checkManifestCase(entry: ManifestCase) {
  const p = await manifestPivot(entry)
  const view = initialViewState(p.def, p.cache), axes = okAxes(p, view)
  const saved = savedPivotLayout(p.def)!
  expect(axes.rows).toEqual(saved.rows)
  expect(axes.cols).toEqual(saved.cols)
  const layout = placeCells(p.def, view, p.cache, axes, blankGrid, { date1904: p.date1904 })
  expect(liveDefinition(p.def, view, axes, layout).location).toEqual(p.def.location)
  expect(labelDiffs(p, layout)).toEqual([])
  // Merged labels (mergeItem): the layout's merges are exactly the saved sheet's merged ranges in the pivot.
  expect([...layout.merges].sort(byRect)).toEqual(savedMerges(p))
}

/** Ruling S2-AA: task-17 manifest cases the engine still refuses, each with its exact refusal and why Excel's file
 * does not settle it. Empty after fix round B1: every task-17 case computes. */
const TASK17_REFUSED: Record<string, { message: string; why: string }> = {}
/** Task-17 cases the engine refuses, each with its exact refusal and why Excel's file does not settle it. */
async function checkManifestStatus(entry: ManifestCase) {
  const p = await manifestPivot(entry)
  const refused = TASK17_REFUSED[entry.case]
  const result = checkSavedLayout(p.store, p.cache, p.def, sheetCellReader(p.sheet), { date1904: p.date1904 })
  if (refused) {
    const computed = computePivot(p.store, p.cache, p.def, initialViewState(p.def, p.cache), { date1904: p.date1904 })
    expect(computed.ok ? undefined : computed.diagnostic.message).toBe(refused.message)
    expect(result.status).toBe('unverified')
    return
  }
  await checkManifestCase(entry)
  expect(result).toEqual({ status: 'ok' })
}

describe('Task 15 Excel oracle: manifest cases with task 17', () => {
  if (!manifest) {
    it.skip('SKIPPED: tests/fixtures/xlsx-pivots/engine/manifest.json is absent (Task 15 oracle not built); task-17 cases not run', () => {})
    return
  }
  if (!task17.length) {
    it.skip('SKIPPED: the manifest has no runnable task-17 case (all pending)', () => {})
    return
  }
  it.each(task17.map(entry => [entry.case, entry] as const))('%s', async (_name, entry) => checkManifestStatus(entry))
  it('the refusal map names task-17 cases only', () => {
    for (const name of Object.keys(TASK17_REFUSED)) expect(task17.map(entry => entry.case)).toContain(name)
  })
})

describe('cross-field visibility: an item left without records by another field\'s hidden item or page selection follows the measurement', () => {
  const crossField = task17.filter(entry => entry.features.some(feature => feature.includes('cross-field')))
  if (!manifest) {
    it.skip('SKIPPED: manifest.json absent (Task 15): the cross-field measurement cases are not run', () => {})
    return
  }
  if (!crossField.length) {
    it.skip('SKIPPED: the manifest has no runnable cross-field task-17 case (all pending)', () => {})
    return
  }
  it.each(crossField.map(entry => [entry.case, entry] as const))('%s', async (_name, entry) => checkManifestCase(entry))
})
