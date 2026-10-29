// @vitest-environment node
/**
 * Pivot display (spec §6.5–§6.6, §7.6, §7.10; Task 10): saved layout → style
 * roles, `formats` overrides, compact indent, and the pivot buttons. Every
 * expectation below cites the Excel 16.114 reference it is read from
 * (tests/fixtures/xlsx-pivots/refs/*.png, docs/validation/xlsx-pivots.md).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test } from 'vitest'
import { parseXml } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, layoutCellBody, textMeasurer } from '../src/xlsx/render'
import { getPaintables } from '../src/render/paint'
import { ensureTableStyleEngine, regionStylesFor } from '../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { orderRoles } from '../src/xlsx/tableStyles/resolve'
import type { RoleRef, TableStyleElementType } from '../src/xlsx/tableStyles/model'
import { savedPivotRoles } from '../src/xlsx/pivots/savedLayout'
import { dropdownBox, pivotButtons, type ButtonPlacement } from '../src/xlsx/pivots/buttons'
import type { PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

const fixture = (name: string) => readFileSync(fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}.xlsx`, import.meta.url)))
const precedenceColors = JSON.parse(readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/recipes/pivot-precedence-colors.json', import.meta.url)), 'utf8')) as Record<string, string>
async function load(name: string): Promise<{ doc: XlsxDocument; sheet: XlsxSheet; def: PivotTableDefinition }> {
  const doc = await parseXlsx(await OfficePackage.load(fixture(name)))
  const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)
  if (!sheet) throw new Error(`${name}: no pivot table definitions on any sheet`)
  await ensureTableStyleEngine(doc.sheets)
  return { doc, sheet, def: sheet.pivotTables![0]! }
}
/** Re-resolve the sheet's region layer after a test edits its definition or styles. */
function relayer(sheet: XlsxSheet) {
  sheet.regionStyles = undefined
  sheet.regionStylesResolved = false
  return regionStylesFor(sheet)
}
const element = (role: RoleRef): TableStyleElementType => typeof role === 'string' ? role : role.element
const winner = (def: PivotTableDefinition, row: number, col: number) => orderRoles(savedPivotRoles(def)!.roleAt(row, col), 'pivot').map(element).at(-1)
const elements = (def: PivotTableDefinition, row: number, col: number) => orderRoles(savedPivotRoles(def)!.roleAt(row, col), 'pivot').map(element)
const band = (def: PivotTableDefinition, row: number, col: number, type: TableStyleElementType) => {
  const role = savedPivotRoles(def)!.roleAt(row, col).find(candidate => element(candidate) === type)
  return role && typeof role !== 'string' ? role.edges : role
}
/** A1-style address → zero-based coordinates. */
const at = (address: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(address)!
  return { row: Number(m[2]) - 1, col: [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1 }
}
const PX_PER_PT = 96 / 72
const pixel = (canvas: ReturnType<typeof createCanvas>, x: number, y: number) => {
  const [r, g, b] = canvas.getContext('2d').getImageData(Math.floor(x), Math.floor(y), 1, 1).data
  return [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()
}

describe('savedPivotRoles: the standard compact pivot (refs/measure-precedence.png)', () => {
  // Winning element per cell, read from refs/measure-precedence.png: custom
  // style OfficeViewPrecedence gives every element its own fill
  // (recipes/pivot-precedence-colors.json), all style-info flags on. The grid
  // is A3:I22, the pivot's saved location.
  const K: Record<string, TableStyleElementType> = {
    FH: 'firstHeaderCell', HR: 'headerRow', C1: 'firstColumnSubheading', C2: 'secondColumnSubheading', LC: 'lastColumn',
    R1: 'firstRowSubheading', FC: 'firstColumn', S1: 'firstRowStripe', S2: 'secondRowStripe', TR: 'totalRow',
  }
  const subheading = 'R1 R1 R1 R1 R1 R1 R1 LC LC'
  const leaf = (stripe: string) => `FC ${Array(6).fill(stripe).join(' ')} LC LC`
  const block = [subheading, leaf('S2'), leaf('S1'), leaf('S2')]
  const grid = [
    'FH HR HR HR HR HR HR HR HR', // row 3: "Column Labels"; headerRow over lastColumn (H3:I3)
    'FH C1 C1 C1 C1 C1 C1 LC LC', // row 4: year items
    'HR C2 C2 C2 C2 C2 C2 LC LC', // row 5: "Row Labels" + Values items
    ...block, ...block, ...block, ...block, // rows 6–21: East, North, South, West and their products
    'TR TR TR TR TR TR TR TR TR', // row 22: Grand Total
  ].map(line => line.split(' ').map(code => K[code]))

  test('the winning role of every cell matches the reference', async () => {
    const { def } = await load('measure-precedence')
    const roles = savedPivotRoles(def)!
    expect(roles.rect).toEqual({ minRow: 2, minCol: 0, maxRow: 21, maxCol: 8 })
    const actual = grid.map((line, r) => line.map((_, c) => winner(def, 2 + r, c)))
    expect(actual).toEqual(grid)
  })

  test('the resolved fill of every cell is the reference colour', async () => {
    const { sheet } = await load('measure-precedence')
    const layer = regionStylesFor(sheet)!
    grid.forEach((line, r) => line.forEach((type, c) => {
      expect(layer.styleAt(2 + r, c, undefined)?.fillColor?.toUpperCase(), `${type} at row ${3 + r} col ${c}`).toBe(`FF${precedenceColors[type]}`)
    }))
  })

  test('exact roles of representative cells, ascending precedence', async () => {
    const { def } = await load('measure-precedence')
    const cell = (address: string) => { const { row, col } = at(address); return elements(def, row, col) }
    expect(cell('A3')).toEqual(['wholeTable', 'headerRow', 'firstHeaderCell'])
    // Ruling AE: the top header row never carries lastColumn (H3 shows headerRow).
    expect(cell('H3')).toEqual(['wholeTable', 'headerRow'])
    expect(cell('B4')).toEqual(['wholeTable', 'headerRow', 'firstColumnSubheading'])
    expect(cell('H4')).toEqual(['wholeTable', 'headerRow', 'lastColumn'])
    // A5 ("Row Labels") is headerRow only: firstHeaderCell stops above it.
    expect(cell('A5')).toEqual(['wholeTable', 'headerRow'])
    expect(cell('B5')).toEqual(['wholeTable', 'headerRow', 'secondColumnSubheading'])
    // Subheading rows count toward the stripe alternation (East = stripe 1).
    expect(cell('A6')).toEqual(['wholeTable', 'firstRowStripe', 'firstColumn', 'firstRowSubheading'])
    expect(cell('B6')).toEqual(['wholeTable', 'firstColumnStripe', 'firstRowStripe', 'firstRowSubheading'])
    expect(cell('H6')).toEqual(['wholeTable', 'firstRowStripe', 'firstRowSubheading', 'lastColumn'])
    expect(cell('A7')).toEqual(['wholeTable', 'secondRowStripe', 'firstColumn'])
    expect(cell('C7')).toEqual(['wholeTable', 'secondColumnStripe', 'secondRowStripe'])
    // The grand-total row carries firstColumn and lastColumn under totalRow.
    expect(cell('A22')).toEqual(['wholeTable', 'firstColumn', 'totalRow'])
    expect(cell('H22')).toEqual(['wholeTable', 'lastColumn', 'totalRow'])
  })

  test('roles carry the bands of their elements', async () => {
    const { def } = await load('measure-precedence')
    const edges = (address: string, type: TableStyleElementType) => { const { row, col } = at(address); return band(def, row, col, type) }
    expect(edges('A3', 'wholeTable')).toEqual({ top: true, bottom: false, left: true, right: false })
    expect(edges('I22', 'wholeTable')).toEqual({ top: false, bottom: true, left: false, right: true })
    expect(edges('B5', 'headerRow')).toEqual({ top: false, bottom: true, left: false, right: false })
    // lastColumn: the two grand-total columns, from the column-item rows down.
    expect(edges('H4', 'lastColumn')).toEqual({ top: true, bottom: false, left: true, right: false })
    expect(edges('I22', 'lastColumn')).toEqual({ top: false, bottom: true, left: false, right: true })
    // A row subheading spans its row; a column subheading its level row of item columns.
    expect(edges('A6', 'firstRowSubheading')).toEqual({ top: true, bottom: true, left: true, right: false })
    expect(edges('G4', 'firstColumnSubheading')).toEqual({ top: true, bottom: true, left: false, right: true })
    expect(edges('A7', 'firstColumn')).toEqual({ top: false, bottom: false, left: true, right: true })
  })
})

describe('savedPivotRoles: subtotals, levels and forms', () => {
  test('a bottom subtotal row is firstSubtotalRow, one band per line (refs/pivot-subtotals-bottom.png)', async () => {
    const { def, sheet } = await load('pivot-subtotals-bottom')
    // Row 10 "East 合計" is bold with lines above and below across the table.
    expect(winner(def, 9, 0)).toBe('firstSubtotalRow')
    expect(winner(def, 9, 4)).toBe('firstSubtotalRow')
    expect(band(def, 9, 0, 'firstSubtotalRow')).toEqual({ top: true, bottom: true, left: true, right: false })
    const style = regionStylesFor(sheet)!.styleAt(9, 3, undefined)!
    expect(style.bold).toBe(true)
    expect(style.borders).toMatchObject({ top: 'thin', bottom: 'thin' })
    expect(style.borderColors?.top?.toUpperCase()).toBe('FF156082')
    expect(style.borderColors?.bottom?.toUpperCase()).toBe('FF156082')
  })

  test('two adjacent subtotal lines are two bands (refs/pivot-values-on-rows.png: lines at y 210, 225 and 240)', async () => {
    const { def, sheet } = await load('pivot-values-on-rows')
    // Rows 15–16: "East Sum of Sales" / "East Sum of Units".
    for (const row of [14, 15]) {
      expect(band(def, row, 2, 'firstSubtotalRow')).toEqual({ top: true, bottom: true, left: false, right: false })
      expect(regionStylesFor(sheet)!.styleAt(row, 2, undefined)!.borders).toMatchObject({ top: 'thin', bottom: 'thin' })
    }
  })

  test('row subheadings follow the item level (refs/pivot-values-on-rows.png: East and Apples bold)', async () => {
    const { def } = await load('pivot-values-on-rows')
    expect(winner(def, at('A5').row, 0)).toBe('firstRowSubheading') // East
    expect(winner(def, at('A6').row, 0)).toBe('secondRowSubheading') // Apples: its children are the Values items
    expect(elements(def, at('A7').row, 0)).not.toContain('thirdRowSubheading') // "Sum of Sales": a leaf
    expect(elements(def, at('A7').row, 0).some(type => type.endsWith('RowSubheading'))).toBe(false)
  })

  test('second and third subtotal rows (ECMA-376 §18.8.40 order; no Excel reference has them)', async () => {
    const { def } = await load('pivot-values-on-rows')
    // Rows Region › Product › Values: give the first three lines subtotal types at levels 0, 1 and 2.
    const synthetic: PivotTableDefinition = structuredClone(def)
    synthetic.rowItems![0] = { type: 'default', r: 0, i: 0, x: [0] }
    synthetic.rowItems![1] = { type: 'default', r: 1, i: 0, x: [0] }
    synthetic.rowItems![2] = { type: 'sum', r: 2, i: 0, x: [0] }
    expect(winner(synthetic, 4, 1)).toBe('firstSubtotalRow')
    expect(winner(synthetic, 5, 1)).toBe('secondSubtotalRow')
    expect(winner(synthetic, 6, 1)).toBe('thirdSubtotalRow')
  })

  test('tabular form: a parent item is a subheading in its own label column only (refs/pivot-layout-tabular.png)', async () => {
    const { def } = await load('pivot-layout-tabular')
    // "East" (A6) is bold with the subheading's light line under A8; B6 "Apples" is plain.
    for (const row of [5, 6, 7]) expect(elements(def, row, 0)).toContain('firstRowSubheading')
    expect(band(def, 5, 0, 'firstRowSubheading')).toEqual({ top: true, bottom: false, left: true, right: true })
    expect(band(def, 7, 0, 'firstRowSubheading')).toEqual({ top: false, bottom: true, left: true, right: true })
    expect(elements(def, 5, 1)).not.toContain('firstRowSubheading')
    expect(elements(def, 5, 4)).not.toContain('firstRowSubheading')
    expect(winner(def, 8, 0)).toBe('firstSubtotalRow') // "East 合計"
  })

  test('tabular form: the parent label\'s bottom edge is kept above its subtotal line (refs/pivot-layout-tabular.png)', async () => {
    const { doc, sheet, def } = await load('pivot-layout-tabular')
    // The A8/A9 edge is #44B3E1 (the subheading's line) in column A and #156082
    // (the subtotal row's top line) from column B on: the subtotal's top is no outer edge in A.
    expect(band(def, 8, 0, 'firstSubtotalRow')).toEqual({ top: false, bottom: true, left: true, right: false })
    expect(band(def, 8, 1, 'firstSubtotalRow')).toEqual({ top: true, bottom: true, left: false, right: false })
    const index = doc.sheets.indexOf(sheet)
    const pages = await getPaintables(doc)
    const canvas = createCanvas(pages[index]!.spec.widthPx, pages[index]!.spec.heightPx)
    try { pages[index]!.paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D) } finally { pages.dispose() }
    const m = computeMetrics(sheet)
    const edge = m.rowHeightsPx.slice(0, 8).reduce((a, b) => a + b, 0)
    const mid = (col: number) => m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0) + m.colWidthsPx[col]! / 2
    // Theme tints resolve within a few RGB levels of Excel's (#46B1E1 here for #44B3E1).
    const near = (actual: string, expected: string) => [0, 2, 4].every(i => Math.abs(parseInt(actual.slice(i, i + 2), 16) - parseInt(expected.slice(i, i + 2), 16)) <= 3)
    for (const [col, colour] of [[0, '44B3E1'], [1, '156082'], [4, '156082']] as const) {
      const actual = pixel(canvas, mid(col), edge)
      expect(near(actual, colour), `col ${col}: ${actual} vs ${colour}`).toBe(true)
    }
  })

  test('outline form: a parent item row is a whole-row subheading (refs/pivot-layout-outline.png)', async () => {
    const { def } = await load('pivot-layout-outline')
    for (const col of [0, 1, 5, 9]) expect(winner(def, 5, col), `col ${col}`).toBe(col >= 8 ? 'lastColumn' : 'firstRowSubheading')
  })

  test('page fields sit above the location (refs/pivot-page-field.png: A1 and B1 shaded)', async () => {
    const { def, sheet } = await load('pivot-page-field')
    const roles = savedPivotRoles(def)!
    expect(roles.rect).toEqual({ minRow: 2, minCol: 0, maxRow: 9, maxCol: 8 })
    expect(roles.rects).toContainEqual({ minRow: 0, minCol: 0, maxRow: 0, maxCol: 1 })
    expect(roles.roleAt(0, 0).map(element)).toEqual(['pageFieldLabels'])
    expect(roles.roleAt(0, 1).map(element)).toEqual(['pageFieldValues'])
    expect(roles.roleAt(1, 0)).toEqual([])
    expect(regionStylesFor(sheet)!.styleAt(0, 1, undefined)?.fillColor).toBeDefined()
  })
})

describe('fail closed (§7.6, Ruling AC)', () => {
  test('no saved rowItems: own s= only and pivot-layout-unavailable', async () => {
    const { sheet, def } = await load('measure-precedence')
    delete def.rowItems
    expect(savedPivotRoles(def)).toBeUndefined()
    const own = { numFmtId: 0, bold: true }
    expect(relayer(sheet)?.styleAt(9, 2, own) ?? own).toBe(own)
    expect(sheet.diagnostics?.some(d => d.feature === 'pivot-layout-unavailable')).toBe(true)
  })

  test('an invalid index: never laid out', async () => {
    const { sheet, def } = await load('measure-precedence')
    def.unsupported = [...def.unsupported, 'invalid-index:rowItems[3].x[0]']
    expect(savedPivotRoles(def)).toBeUndefined()
    expect(relayer(sheet)?.styleAt(9, 2, undefined)?.fillColor).toBeUndefined()
    expect(sheet.diagnostics?.some(d => d.feature === 'pivot-layout-unavailable')).toBe(true)
  })

  test('saved items that do not fill the location are not laid out', async () => {
    const { def } = await load('measure-precedence')
    def.rowItems = def.rowItems!.slice(1)
    expect(savedPivotRoles(def)).toBeUndefined()
  })

  test('other unsupported constructs keep the saved display and its style', async () => {
    const { sheet, def } = await load('measure-precedence')
    def.unsupported = [...def.unsupported, 'pivotTableDefinition/pivotHierarchies']
    expect(winner(def, 9, 0)).toBe('firstRowSubheading')
    expect(relayer(sheet)!.styleAt(9, 0, undefined)?.fillColor?.toUpperCase()).toBe(`FF${precedenceColors.firstRowSubheading}`)
  })

  test('pivot conditional formats are reported and skipped (no conditional-format engine)', async () => {
    const { sheet, def } = await load('measure-precedence')
    def.conditionalFormats = [{ scope: 'selection', type: 'none', priority: 1, pivotAreas: [] }]
    expect(relayer(sheet)!.styleAt(9, 0, undefined)?.fillColor?.toUpperCase()).toBe(`FF${precedenceColors.firstRowSubheading}`)
    expect(sheet.diagnostics?.some(d => d.feature === 'pivot-conditional-format')).toBe(true)
  })
})

describe('formats overrides (refs/pivot-formats-override.png)', () => {
  // The fixture's format: Sum of Sales × East × Apples × 2022 → B7, fill FFE699.
  // Excel also writes that fill into B7's own s=, so each test strips B7's own
  // style to keep the pivotArea match under test.
  const STYLES = `<styleSheet><dxfs>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFE699"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FF0000FF"/></patternFill></fill></dxf>
    </dxfs><tableStyles><tableStyle name="OVFormats" table="0" pivot="1">
    <tableStyleElement type="wholeTable" dxfId="1"/>
    <tableStyleElement type="firstRowStripe" dxfId="2"/>
    <tableStyleElement type="secondRowStripe" dxfId="2"/>
    </tableStyle></tableStyles></styleSheet>`

  test('the format lands on the cells its pivotArea names', async () => {
    const { sheet } = await load('pivot-formats-override')
    const layer = regionStylesFor(sheet)!
    expect(layer.styleAt(6, 1, undefined)?.fillColor).toBe('FFFFE699') // B7
    expect(layer.styleAt(7, 1, undefined)?.fillColor).toBeUndefined() // B8 (Bananas)
    expect(layer.styleAt(6, 2, undefined)?.fillColor).toBeUndefined() // C7 (Sum of Units)
    expect(layer.styleAt(5, 1, undefined)?.fillColor).toBeUndefined() // B6 (East subtotal)
    expect(layer.styleAt(6, 0, undefined)?.fillColor).toBeUndefined() // A7 (the label)
  })

  test('a format wins over every style element; the cell\'s own fill wins over the format', async () => {
    const { sheet, def } = await load('pivot-formats-override')
    sheet.tableStyleSource = { root: parseXml(STYLES), colors: {} }
    def.styleInfo = { ...def.styleInfo!, name: 'OVFormats', showRowStripes: true }
    const layer = relayer(sheet)!
    expect(layer.styleAt(7, 1, undefined)?.fillColor).toBe('FF0000FF') // B8: the stripe over wholeTable
    expect(layer.styleAt(6, 1, undefined)?.fillColor).toBe('FFFFE699') // B7: the format over the stripe
    // measure-cell-override / Ruling P: an own fill draws over the pivot style.
    expect(layer.styleAt(6, 1, { numFmtId: 0, fillColor: 'FF00FF00' })?.fillColor).toBe('FF00FF00')
  })
})

describe('pivot buttons (§7.10)', () => {
  const sorted = (list: ButtonPlacement[]) => [...list].sort((a, b) => a.row - b.row || a.col - b.col || a.kind.localeCompare(b.kind))

  test('screen: Row Labels, Column Labels and page-field dropdowns (refs/measure-field-buttons.png)', async () => {
    const { def } = await load('pivot-page-field')
    // The window screenshot shows dropdowns at B1 (page field Product), B3 (Column Labels) and A5 (Row Labels).
    expect(sorted(pivotButtons(def, 'screen'))).toEqual([
      { row: 0, col: 1, kind: 'page', field: 1 },
      { row: 2, col: 1, kind: 'filter', field: 2 },
      { row: 4, col: 0, kind: 'filter', field: 0 },
    ])
  })

  test('screen: outline and tabular forms have one dropdown per field header plus the ± buttons', async () => {
    for (const name of ['pivot-layout-outline', 'pivot-layout-tabular']) {
      const { def } = await load(name)
      // Dropdowns on Year (C3), Region (A5) and Product (B5), none on Values (D3)
      // although its xf carries pivotButton="1" (Ruling AF2,
      // t10-excel/shot-pivot-layout-outline.png); ⊖ on East, North, South and
      // West in column A (refs/<name>.png).
      expect(sorted(pivotButtons(def, 'screen')), name).toEqual([
        { row: 2, col: 2, kind: 'filter', field: 2 },
        { row: 4, col: 0, kind: 'filter', field: 0 },
        { row: 4, col: 1, kind: 'filter', field: 1 },
        { row: 5, col: 0, kind: 'collapse', field: 0, item: 0 },
        { row: 9, col: 0, kind: 'collapse', field: 0, item: 1 },
        { row: 13, col: 0, kind: 'collapse', field: 0, item: 2 },
        { row: 17, col: 0, kind: 'collapse', field: 0, item: 3 },
      ])
    }
  })

  test('screen: a collapsed item gets the expand (+) button (refs/pivot-collapsed-item.png)', async () => {
    const { def } = await load('pivot-collapsed-item')
    expect(sorted(pivotButtons(def, 'screen')).filter(b => b.kind !== 'filter')).toEqual([
      { row: 5, col: 0, kind: 'expand', field: 0, item: 0 },
      { row: 6, col: 0, kind: 'collapse', field: 0, item: 1 },
      { row: 10, col: 0, kind: 'collapse', field: 0, item: 2 },
      { row: 14, col: 0, kind: 'collapse', field: 0, item: 3 },
    ])
  })

  test('showDataDropDown off hides every dropdown and keeps the ± buttons; showHeaders hides only the field-header ones', async () => {
    const page = await load('pivot-page-field')
    page.def.showDataDropDown = false
    expect(pivotButtons(page.def, 'screen')).toEqual([])
    const outline = await load('pivot-layout-outline')
    outline.def.showDataDropDown = false
    expect(sorted(pivotButtons(outline.def, 'screen'))).toEqual([
      { row: 5, col: 0, kind: 'collapse', field: 0, item: 0 },
      { row: 9, col: 0, kind: 'collapse', field: 0, item: 1 },
      { row: 13, col: 0, kind: 'collapse', field: 0, item: 2 },
      { row: 17, col: 0, kind: 'collapse', field: 0, item: 3 },
    ])
    const headers = await load('pivot-page-field')
    headers.def.showHeaders = false
    expect(pivotButtons(headers.def, 'screen')).toEqual([{ row: 0, col: 1, kind: 'page', field: 1 }])
  })

  test('the dropdown box is 15 pt wide and as tall as the row up to 15 pt', () => {
    const w = Math.round(15 * PX_PER_PT)
    // A 15-pt row (20 px): 15 × 15 (t10-excel/shot-pivot-layout-outline.png, border y 347–361).
    expect(dropdownBox({ x: 100, y: 40, width: 80, height: 20 })).toEqual({ x: 180 - w, y: 40, width: w, height: 20 })
    // A shorter row: the box takes the row's height, still flush right and bottom.
    expect(dropdownBox({ x: 100, y: 40, width: 80, height: 16 })).toEqual({ x: 180 - w, y: 40, width: w, height: 16 })
    // A taller row: still 15 pt, at the bottom.
    expect(dropdownBox({ x: 100, y: 40, width: 80, height: 40 })).toEqual({ x: 180 - w, y: 60, width: w, height: 20 })
  })

  test('print: no dropdowns; ± only when printDrill is set', async () => {
    const { def } = await load('pivot-collapsed-item')
    expect(pivotButtons(def, 'print')).toEqual([])
    const page = await load('pivot-page-field')
    expect(pivotButtons(page.def, 'print')).toEqual([])
    def.printDrill = true
    expect(sorted(pivotButtons(def, 'print'))).toEqual(sorted(pivotButtons(def, 'screen')).filter(b => b.kind !== 'filter' && b.kind !== 'page'))
    def.showDrill = false
    expect(pivotButtons(def, 'screen').filter(b => b.kind === 'expand' || b.kind === 'collapse')).toEqual([])
  })

  test('buttons follow print in page mode and screen in the viewer (§7.10)', async () => {
    const { doc, sheet } = await load('pivot-collapsed-item')
    const m = computeMetrics(sheet)
    const x = (col: number) => m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0)
    const y = (row: number) => m.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0)
    // The Column Labels dropdown (B3): a 15×14 pt box flush with the cell's right
    // and bottom edges, #738A93 border, #FDFDFD fill (measure-field-buttons).
    const dropLeft = x(2) - Math.round(15 * PX_PER_PT), dropMid = y(3) - Math.round(7 * PX_PER_PT)
    // The East + button (A6): a 7×7 pt box at x 6–12 pt, vertically centred.
    const plusLeft = x(0) + Math.round(6 * PX_PER_PT)
    const plusTop = y(5) + Math.round((m.rowHeightsPx[5]! - Math.round(7 * PX_PER_PT)) / 2)

    const pages = await getPaintables(doc)
    const page = pages[doc.sheets.indexOf(sheet)]!
    const still = createCanvas(page.spec.widthPx, page.spec.heightPx)
    try { page.paint(still.getContext('2d') as unknown as CanvasRenderingContext2D) } finally { pages.dispose() }

    const geometry = buildSheetGeometry(sheet)
    expect(geometry.kind).toBe('ready')
    const live = createCanvas(400, 200)
    const result = renderSheetViewport(sheet, live.getContext('2d') as never, geometry, { box: { x: 0, y: 0, width: 400, height: 200 }, zoom: 1, deviceScale: 1 }, {}, {
      owner: { docId: 'pivot-buttons', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }, cache: new LayoutCache(),
    })
    expect(result.kind).toBe('detail')

    expect(pixel(live, dropLeft, dropMid)).toBe('738A93')
    // 15 × 15 pt in a 15-pt row: the top border lies on the row's first pixel (shot-pivot-layout-outline).
    expect(pixel(live, dropLeft + 10, y(2))).toBe('738A93')
    expect(pixel(live, dropLeft + 10, y(2) + 2)).toBe('FDFDFD')
    expect(pixel(live, dropLeft + 3, dropMid - 4)).toBe('FDFDFD')
    // The arrow: rows 7/5/3/1 pt wide from 6 pt below the top border, ≈ #828282 at its
    // centre (shot-pivot-layout-outline: rows 353–356 under the border at 347). The box
    // is 15 pt wide, so the arrow's centre is 7.5 pt in; background above and below it.
    const arrowX = dropLeft + Math.round(15 * PX_PER_PT) / 2, arrowTop = y(2) + Math.round(6 * PX_PER_PT)
    const ink = pixel(live, arrowX, arrowTop)
    expect([0, 2, 4].every(i => Math.abs(parseInt(ink.slice(i, i + 2), 16) - 0x82) <= 16), `arrow top row ${ink}`).toBe(true)
    expect(pixel(live, arrowX, arrowTop - 1)).toBe('FDFDFD')
    expect(pixel(live, arrowX, arrowTop + Math.ceil(4 * PX_PER_PT) + 1)).toBe('FDFDFD')
    expect(pixel(still, dropLeft, dropMid)).not.toBe('738A93')
    expect(pixel(live, plusLeft, plusTop + 3)).toBe('AAB5C9')
    expect(pixel(still, plusLeft, plusTop + 3)).not.toBe('AAB5C9')
  })
})

describe('compact indent (refs/pivot-values-on-rows.png: text at 15 + 10 × level pt)', () => {
  test('screen: data labels start at 15 + 10 × level pt; subtotal, grand and header labels do not move', async () => {
    const { sheet } = await load('pivot-values-on-rows')
    const layer = regionStylesFor(sheet)!
    const ctx = createCanvas(200, 40).getContext('2d') as unknown as CanvasRenderingContext2D
    const measured = textMeasurer(ctx)
    const cellAt = (row: number) => sheet.rows.flatMap(r => r.cells).find(c => c.row === row && c.col === 0)!
    const textStart = (row: number, mode: 'screen' | 'print') => {
      const cell = cellAt(row)
      const body = layoutCellBody(ctx, sheet, { ...cell, style: layer.styleAt(row, 0, cell.style, mode) }, 0, 0, 160, 20, measured)
      const text = body.text!
      return text.kind === 'plain' ? text.tx : text.layout.lines[0]!.segments[0]!.transform!.e
    }
    expect(textStart(at('A5').row, 'screen')).toBeCloseTo(15 * PX_PER_PT, 1) // East, level 0
    expect(textStart(at('A6').row, 'screen')).toBeCloseTo(25 * PX_PER_PT, 1) // Apples, level 1
    expect(textStart(at('A7').row, 'screen')).toBeCloseTo(35 * PX_PER_PT, 1) // Sum of Sales, level 2
    // "East Sum of Sales" (subtotal), "Row Labels" and the grand total keep the cell's own placement.
    for (const address of ['A15', 'A4', 'A53']) {
      const { row } = at(address)
      expect(layer.styleAt(row, 0, cellAt(row).style, 'screen')?.indentPx, address).toBeUndefined()
    }
    // Print draws no ± buttons: labels keep the cell's own indent (level × MDW).
    expect(layer.styleAt(at('A6').row, 0, cellAt(at('A6').row).style, 'print')?.indentPx).toBeUndefined()
  })
})
