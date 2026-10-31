// @vitest-environment node
/**
 * Plan Task 29 (spec §7.1, §7.8, §7.9; rulings S3-U … S3-Y): the overlap rule, two filter families, the sheet
 * sort/filter rules beside and over a pivot, and column autofit after a change — each against the Excel recipes
 * measured 2026-11-01 (docs/validation/xlsx-pivots.md, "Owner recipes", "Autofit after a change").
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createResidentPivotClient, pendingPivotFits, valueFilter } from '../src/components/xlsx/pivot/pivotClient'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import { loadPivotParts } from './helpers/xlsx-pivots-load'
import JSZip from 'jszip'
import { SheetSizes } from '../src/xlsx/view/axisSizes'
import { macDigitWidth, PIVOT_FIT_INDENT_PT, PIVOT_FIT_TOLERANCE, pivotAutofit, pivotColumnFits } from '../src/xlsx/pivots/session/autofit'
import { calibriTextPt } from './helpers/calibri-model'
import { ensureTableStyleEngine, regionStylesFor } from '../src/xlsx/tableStyleLoader'
import { cellDisplayText } from '../src/xlsx/render'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { guardFilter, guardSort, PIVOT_SHEET_CHANGE_MESSAGE } from '../src/xlsx/view/operationGuard'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const file = (name: string) => resolve(FIXTURES, `${name}.xlsx`)
const signal = () => new AbortController().signal
async function load(name: string): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file(name)))))
}
const pivotSheet = (doc: XlsxDocument) => doc.sheets.findIndex(sheet => sheet.pivotTables?.length)
function client(doc: XlsxDocument) {
  const installs: any[] = []
  return { installs, client: createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => installs.push(commit) }) }
}
/** Puts a cell on the sheet (a value, or a style alone for a formatting-only cell). */
function putCell(sheet: XlsxSheet, row: number, col: number, value: string | null) {
  let line = sheet.rows.find(entry => entry.index === row)
  if (!line) { line = { index: row, cells: [] } as unknown as XlsxSheet['rows'][number]; sheet.rows.push(line); sheet.rows.sort((a, b) => a.index - b.index) }
  line.cells.push({ ref: `A${row + 1}`, row, col, value, styleIndex: 1, style: { numFmtId: 0, fillColor: 'FFFFFF00' } } as never)
}
const n15 = (v: number) => Number(v.toPrecision(15))
/** The layout's cells against a saved sheet over the saved pivot's rectangles (values, numbers to 15 digits). */
function sameCells(layout: { cells: readonly { row: number; col: number; value: { kind: string; value?: unknown } }[] }, saved: XlsxSheet, rects: readonly { minRow: number; maxRow: number; minCol: number; maxCol: number }[]): string[] {
  const read = sheetCellReader(saved), ours = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell.value]))
  const out: string[] = []
  for (const rect of rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
    const theirs = read(row, col), mine = ours.get(`${row},${col}`)
    const a = mine && mine.kind !== 'empty' ? mine.kind === 'number' ? n15(mine.value as number) : String(mine.value) : undefined
    const b = theirs ? theirs.kind === 'number' ? n15(theirs.value) : String(theirs.value) : undefined
    // Implicit captions are localised by the session; compared for presence.
    if ((a === undefined) !== (b === undefined) || (typeof a === 'number' && a !== b)) out.push(`R${row + 1}C${col + 1}: ${a} vs ${b}`)
  }
  return out
}

describe('S3-Y: two filter families with multipleFieldFilters="0"', () => {
  it('a filter of the other family replaces the field\'s filter, as Excel saved it (owner-filter-two-families)', async () => {
    const doc = await load('engine/filter-label-begins-with'), sheet = pivotSheet(doc)
    const def = doc.sheets[sheet]!.pivotTables![0]!
    expect(def.multipleFieldFilters).toBe(false)
    const { client: c } = client(doc)
    const opened = await c.open(sheet, 0, signal())
    if (!opened.ok) throw new Error(JSON.stringify(opened))
    const product = opened.model.fields.find(field => field.name === 'Product')!
    expect(product.filters.map(filter => filter.type)).toEqual(['captionBeginsWith'])
    const applied = await c.apply(sheet, 0, { kind: 'filter', field: product.index, filter: valueFilter(product.index, opened.model.nextFilterId, 0, 'gt', 3000) })
    if (!applied.ok) throw new Error(JSON.stringify(applied))
    expect(applied.model.fields.find(field => field.name === 'Product')!.filters.map(filter => filter.type)).toEqual(['valueGreaterThan'])
    const target = (await loadPivotParts(file('interaction/owner-filter-two-families')))[0]!
    expect(applied.def.location.ref).toBe(target.def.location.ref)
    expect(applied.def.location.ref).toBe('A3:I22')
    const layoutCells = [...savedPivotLayout(applied.def)!.rects]
    const overlay = applied.commit.overlay!
    const live = { cells: layoutCells.flatMap(rect => { const out: any[] = []; for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) { const cell = overlay.cell(row, col); if (cell) out.push({ row, col, value: typeof cell.value === 'number' ? { kind: 'number', value: cell.value } : { kind: 'text', value: cell.value } }) } return out }) }
    expect(sameCells(live, target.sheet, savedPivotLayout(target.def)!.rects)).toEqual([])
    // The grand total Excel saved: 221677.36.
    expect(live.cells.some((cell: any) => cell.value.kind === 'number' && n15(cell.value.value) === 221677.36)).toBe(true)
  })
})

describe('S3-U / S3-V: overlap', () => {
  it('a change into a non-empty cell refuses with the measured message and leaves the pivot as it was', async () => {
    const doc = await load('engine/layout-collapsed-item'), sheet = pivotSheet(doc)
    putCell(doc.sheets[sheet]!, 22, 0, 'blocker')
    const { client: c, installs } = client(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const refused = await c.apply(sheet, 0, { kind: 'collapse', field: 0, item: 0, collapsed: false })
    expect(refused.ok).toBe(false)
    if (!refused.ok) {
      expect(refused.diagnostic.feature).toBe('pivot-overlap')
      expect(refused.diagnostic.message).toBe('This change would overlap cell A23 on Pivot.')
    }
    expect(installs).toHaveLength(0)
  })

  it('formatting-only cells are not occupied: the pivot grows over them (owner-overlap-format)', async () => {
    const doc = await load('engine/layout-collapsed-item'), sheet = pivotSheet(doc)
    putCell(doc.sheets[sheet]!, 22, 0, null)
    const { client: c } = client(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const applied = await c.apply(sheet, 0, { kind: 'collapse', field: 0, item: 0, collapsed: false })
    if (!applied.ok) throw new Error(JSON.stringify(applied))
    const target = (await loadPivotParts(file('interaction/owner-overlap-format')))[0]!
    expect(applied.def.location.ref).toBe(target.def.location.ref)
    // A23 is the pivot's now (its value, its style: the yellow fill is not kept).
    expect(applied.commit.overlay!.cell(22, 0)).toMatchObject({ value: 'Elderberries' })
  })

  it('another pivot\'s cells refuse, naming its first cell (owner-overlap-pivot\'s second pivot)', async () => {
    const doc = await load('engine/layout-collapsed-item'), sheet = pivotSheet(doc)
    const second = (await loadPivotParts(file('interaction/owner-overlap-pivot'))).find(parts => parts.def.name === '樞紐分析表1')!.def
    expect(second.location.ref).toBe('A23:B29')
    doc.sheets[sheet]!.pivotTables = [...doc.sheets[sheet]!.pivotTables!, { ...second, sheet: doc.sheets[sheet]!.name }]
    const { client: c, installs } = client(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const refused = await c.apply(sheet, 0, { kind: 'collapse', field: 0, item: 0, collapsed: false })
    expect(refused.ok).toBe(false)
    if (!refused.ok) expect(refused.diagnostic.message).toBe('This change would overlap cell A23 on Pivot.')
    expect(installs).toHaveLength(0)
  })
})

/** layout-collapsed-item with one more cell in A23 (a value, or a style alone). */
async function withA23(cell: string): Promise<Blob> {
  const zip = await JSZip.loadAsync(readFileSync(file('engine/layout-collapsed-item')))
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', xml.replace('</sheetData>', `<row r="23">${cell}</row></sheetData>`))
  return new Blob([await zip.generateAsync({ type: 'arraybuffer' })])
}
let serial = 0
const id = () => `overlap-${++serial}`
async function pagedApply(blob: Blob): Promise<WorkerReply> {
  const docId = id()
  const imported = await openImportSession({ kind: 'blob', blob }, { docId, profile: 'D' })
  await imported.done
  const source = createPagedWorkbookSource(imported)
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, { yield: async () => {} })
  const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
  await run({ kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
  const sheetId = source.sheets.find(sheet => sheet.name === 'Pivot')!.sheetId
  const opened = await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  expect(opened).toMatchObject({ kind: 'complete' })
  const reply = await run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: { kind: 'collapse', field: 0, item: 0, collapsed: false }, grantBytes: 16 * 1024 * 1024 })
  await run({ kind: 'close', requestId: id(), docId })
  return reply
}

describe('S3-U / S3-V on the paged worker store', () => {
  it('a value refuses naming the cell; a formatting-only cell is grown over', async () => {
    const refused = await pagedApply(await withA23('<c r="A23" t="inlineStr"><is><t>blocker</t></is></c>'))
    expect(refused.kind).toBe('error')
    expect(JSON.stringify(refused)).toContain('This change would overlap cell A23 on Pivot.')
    const grown = await pagedApply(await withA23('<c r="A23" s="1"/>'))
    expect(grown).toMatchObject({ kind: 'complete', result: { kind: 'pivotApply' } })
    if (grown.kind === 'complete' && grown.result.kind === 'pivotApply') expect(grown.result.def.location.ref).toBe('A3:I24')
  })
})

describe('S3-W / S3-X: sheet sort and filter around a pivot (owner-sheet-sort, owner-sheet-filter)', () => {
  // Pivot A3:I24 (rows 2–23, columns 0–8); K5:L10 is the data Excel filtered beside it (rows 4–9, columns 10–11).
  const pivot = { r0: 2, r1: 23, c0: 0, c1: 8 }
  const structure = new StructuralIndex({ stamp: { docId: 'd', sheetId: 's', valueGen: 0 }, completeness: 'complete', missing: [], merges: [], tables: [], spills: [], drawings: [], pivots: [pivot] } as never)
  const range = (r0: number, r1: number, c0: number, c1: number) => ({ r0, r1, c0, c1, headerRows: 1 as const, totalsRows: 0 as const })
  const empty = { iterate: async () => ({ items: [], completeness: 'complete', missing: [] }) } as never
  const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
  const filterOver = (target: ReturnType<typeof range>, hidden: Array<{ r0: number; r1: number }>) => {
    const before = visibilityFromHidden(createIntervalSet([])), after = visibilityFromHidden(createIntervalSet(hidden))
    return guardFilter({ range: target, candidateVisibility: after, currentVisibility: before, structure,
      projectionBefore: createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: before }), projectionAfter: createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: after }) })
  }

  it('S3-W: a sort touching a pivot, or inside it, refuses with Excel\'s alert (in English)', async () => {
    expect(PIVOT_SHEET_CHANGE_MESSAGE).toBe("We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.")
    for (const target of [range(0, 29, 0, 10), range(5, 9, 1, 1)]) {
      expect(await guardSort(target, structure, empty)).toMatchObject({ ok: false, refusal: { reason: 'pivot', message: PIVOT_SHEET_CHANGE_MESSAGE } })
    }
    // The viewer's sort permutes whole display rows: one beside the pivot in its rows would scramble it, and refuses too.
    expect(await guardSort(range(4, 9, 10, 11), structure, empty)).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    expect(await guardSort(range(30, 40, 0, 3), structure, empty)).toEqual({ ok: true })
  })

  it('S3-X: a filter beside a pivot in its rows is allowed and may hide its rows; one intersecting it refuses', async () => {
    // Excel: autoFilter K5:L10, Key = "a" hides sheet rows 6, 8 and 10 (pivot lines East and North included).
    expect(await filterOver(range(4, 9, 10, 11), [{ r0: 5, r1: 5 }, { r0: 7, r1: 7 }, { r0: 9, r1: 9 }])).toEqual({ ok: true })
    for (const target of [range(4, 23, 0, 8), range(4, 9, 8, 11)]) {
      expect(await filterOver(target, [{ r0: 5, r1: 5 }])).toMatchObject({ ok: false, refusal: { reason: 'pivot', message: PIVOT_SHEET_CHANGE_MESSAGE } })
    }
  })
})

describe('SheetSizes records which sizes the user set', () => {
  it('a user resize is user-set; an autofit size is not; reset clears both', () => {
    const sizes = new SheetSizes()
    expect(sizes.set('col', 1, 120)).toBe(true)
    expect(sizes.userSet('col', 1)).toBe(true)
    expect(sizes.set('col', 2, 90, 'autofit')).toBe(true)
    expect(sizes.userSet('col', 2)).toBe(false)
    expect(sizes.get('col', 2)).toBe(90)
    expect(sizes.userSet('row', 1)).toBe(false)
    // A later user resize of an autofit column makes it the user's.
    sizes.set('col', 2, 95)
    expect(sizes.userSet('col', 2)).toBe(true)
    sizes.restore('col', 1, undefined)
    expect(sizes.userSet('col', 1)).toBe(false)
    sizes.reset()
    expect(sizes.userSet('col', 2)).toBe(false)
  })
})

describe('§7.8 autofit after a change', () => {
  const SEEDS: Record<string, string> = { 'engine-seed': 'recipes/engine-seed', 'chart-seed': 'recipes/pivot-chart-seed', 'getpivotdata-seed': 'recipes/getpivotdata-seed', 'refresh-seed': 'recipes/refresh-seed' }
  /** An Excel-saved pivot: its column widths, and its columns' fits under the Calibri-like measure (bold from the
   * pivot style, text as the cell paints it). */
  async function fitted(name: string) {
    const doc = await load(name)
    await ensureTableStyleEngine(doc.sheets)
    const sheet = doc.sheets[pivotSheet(doc)]!, def = sheet.pivotTables![0]!
    const styles = regionStylesFor(sheet)
    const cells = new Map<string, NonNullable<XlsxSheet['rows'][number]['cells'][number]>>()
    for (const line of sheet.rows) for (const cell of line.cells) cells.set(`${line.index},${cell.col}`, cell)
    const rects = savedPivotLayout(def)!.rects
    const inPivot = [...cells.values()].filter(cell => rects.some(rect => cell.row >= rect.minRow && cell.row <= rect.maxRow && cell.col >= rect.minCol && cell.col <= rect.maxCol))
    const measure = { textPt(row: number, col: number) {
      const cell = cells.get(`${row},${col}`)
      if (!cell || cell.value === null || cell.value === '') return undefined
      const style = styles?.styleAt(row, col, cell.style, 'screen') ?? cell.style
      // The cell's own alignment indent, unless the pivot's compact indent replaces it (autofit adds that one).
      const own = (style as { indentPx?: number } | undefined)?.indentPx ? 0 : (style?.indent ?? 0) * PIVOT_FIT_INDENT_PT
      return calibriTextPt(cellDisplayText(sheet, { value: cell.value, style }), Boolean(style?.bold)) + own
    } }
    const widths = new Map<number, number>()
    for (const spec of sheet.cols) for (let col = spec.min; col <= spec.max; col++) if (spec.widthChars !== undefined) widths.set(col, spec.widthChars)
    return { def, widths, fits: pivotColumnFits({ def, cells: inPivot }, measure, macDigitWidth(7)) }
  }
  const manifest = (JSON.parse(readFileSync(resolve(FIXTURES, 'interaction/manifest.json'), 'utf8')) as Array<{ case: string; file: string; seed: string; base?: string; setup?: string }>)
    .filter(entry => !entry.setup && SEEDS[entry.seed])

  it('autofit widths equal Excel\'s saved <col> widths for every interaction case, within one Excel pixel (1/6 char + the 1/256 truncation)', async () => {
    const report: string[] = []
    let columns = 0, changed = 0
    for (const entry of manifest) {
      const start = await fitted(entry.base ?? SEEDS[entry.seed]!)
      const saved = await fitted(entry.file.replace('tests/fixtures/xlsx-pivots/', '').replace('.xlsx', ''))
      const resized = pivotAutofit(start.fits, saved.fits, col => start.widths.get(col), () => false)
      for (const col of new Set([...start.fits.keys(), ...saved.fits.keys()])) {
        columns++
        const excel = saved.widths.get(col), ours = resized.get(col) ?? start.widths.get(col)
        if ((excel !== start.widths.get(col)) !== resized.has(col)) report.push(`${entry.case} col ${col}: Excel ${start.widths.get(col)} → ${excel}, autofit ${resized.has(col) ? `→ ${resized.get(col)}` : 'keeps it'}`)
        else if (excel !== undefined && ours !== undefined && Math.abs(excel - ours) > PIVOT_FIT_TOLERANCE) report.push(`${entry.case} col ${col}: Excel ${excel}, autofit ${ours}`)
        if (resized.has(col)) changed++
      }
    }
    expect(report).toEqual([])
    expect(manifest.length).toBeGreaterThan(40)
    expect(changed).toBeGreaterThan(10)
    expect(columns).toBeGreaterThan(400)
  }, 120_000)

  it('a user-resized column keeps its width', async () => {
    const start = await fitted('recipes/engine-seed'), saved = await fitted('interaction/interaction-collapse-field-compact-top')
    expect(pivotAutofit(start.fits, saved.fits, col => start.widths.get(col), () => false).get(0)).toBeCloseTo(12.83203125, 2)
    expect(pivotAutofit(start.fits, saved.fits, col => start.widths.get(col), col => col === 0).has(0)).toBe(false)
  })
})

describe('§7.8 paged fits waiting for the refreshed display (Task 29 fix round 1)', () => {
  const layout = (name: string) => ({ def: { name } as never, cells: [] })
  const fit = (before: string, after: string) => ({ before: layout(before), after: layout(after) })
  it('two pivots committing before the display refreshes are both fitted, in commit order', () => {
    const pending = pendingPivotFits<string>()
    pending.add('s1', { pivot: 0, fit: fit('a0', 'a1'), styleAt: 'old', styleGen: 'g0' })
    pending.add('s1', { pivot: 1, fit: fit('b0', 'b1'), styleAt: 'old', styleGen: 'g0' })
    expect(pending.take('s1', 'g0')).toEqual([])
    const ready = pending.take('s1', 'g1')
    expect(ready.map(entry => [entry.pivot, (entry.fit.before.def as { name: string }).name, (entry.fit.after.def as { name: string }).name])).toEqual([[0, 'a0', 'a1'], [1, 'b0', 'b1']])
    expect(pending.size).toBe(0)
  })
  it('one pivot committing twice is fitted once, from its first "before" to its last "after"', () => {
    const pending = pendingPivotFits<string>()
    pending.add('s1', { pivot: 0, fit: fit('a0', 'a1'), styleAt: 'first', styleGen: 'g0' })
    pending.add('s1', { pivot: 0, fit: fit('a1', 'a2'), styleAt: 'second', styleGen: 'g0' })
    pending.add('s2', { pivot: 0, fit: fit('c0', 'c1'), styleAt: 'other', styleGen: 'h0' })
    const ready = pending.take('s1', 'g1')
    expect(ready).toHaveLength(1)
    expect(ready[0]).toMatchObject({ pivot: 0, styleAt: 'first' })
    expect([(ready[0]!.fit.before.def as { name: string }).name, (ready[0]!.fit.after.def as { name: string }).name]).toEqual(['a0', 'a2'])
    expect(pending.sheets()).toEqual(['s2'])
    pending.clear()
    expect(pending.size).toBe(0)
  })
})
