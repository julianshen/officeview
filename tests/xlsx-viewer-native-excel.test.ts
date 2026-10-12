/**
 * S13 native Excel checkpoint (AC-STORE-01, AC-SEM-07, AC-REG-01 native legs).
 *
 * `native-corner.xlsx` was written by Microsoft Excel 16 for macOS (zh-TW UI),
 * driven through AppleScript on the user's machine with their approval, and
 * reopened in Excel without a repair prompt (screenshot of the Excel window:
 * docs/validation/native/native-corner-excel-reopen-window.png). The expectations
 * below are the values Excel displayed after reopening, read back with
 * AppleScript `string value`; they are Excel's word, not the viewer's. Both the
 * object (resident) adapter and the paged adapter must agree with each other
 * AND with Excel. A failure here is a viewer/Excel disagreement to report, never
 * an expectation to relax.
 */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cellDisplayText } from '../src/xlsx/render'
import type { XlsxSheet } from '../src/xlsx/types'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import type { WorkbookSource } from '../src/worker/viewerSession'
import type { CellRead, ColumnLanes, SourceStore } from '../src/xlsx/view/contracts'

const FIXTURE = join(__dirname, 'fixtures', 'xlsx-viewer', 'native-corner.xlsx')
const SHA256 = '09104b6d0123eb53c9cd1a7ec900fce780d2c1520abc357f3049dc8631ede82c'
const bytes = new Uint8Array(readFileSync(FIXTURE))

/** "Ünïcødé 漢字 🎉" as Excel stored it: precomposed (NFC) letters and the
 * U+1F389 party popper as the surrogate pair D83C DF89. */
const A6_UNITS = [0xdc, 0x6e, 0xef, 0x63, 0xf8, 0x64, 0xe9, 0x20, 0x6f22, 0x5b57, 0x20, 0xd83c, 0xdf89]

const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const WHOLE = { r0: 0, r1: 1048575, c0: 0, c1: 16383 }

async function bothAdapters(): Promise<{ object: WorkbookSource; paged: WorkbookSource; sheet: XlsxSheet }> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const object = createObjectWorkbookSource(doc, 'native')
  const store = new SourcePageStore({ docId: 'native', backend: memoryPageBackend() })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'native', profile: 'D', store, pageRows: 4 })
  expect((await session.done).kind).toBe('complete')
  return { object, paged: createPagedWorkbookSource(session), sheet: doc.sheets[0] }
}

function cellOf(read: CellRead) {
  expect(read.kind, JSON.stringify(read)).toBe('cell')
  if (read.kind !== 'cell') throw new Error('unreachable')
  return read.cell
}

describe('native Excel checkpoint: native-corner.xlsx (Excel 16 macOS, zh-TW)', () => {
  it('the fixture bytes are the ones Excel wrote (SHA-256 pinned)', () => {
    expect(bytes.byteLength).toBe(10_603)
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(SHA256)
  })

  it('both adapters agree with each other and with Excel on every authored cell', async () => {
    const { object, paged } = await bothAdapters()
    expect(object.sheets).toEqual([{ sheetId: '1', name: '工作表1' }])
    expect(paged.sheets).toEqual(object.sheets)
    const a = object.store('1'), b = paged.store('1')
    // [address, row, col, Excel's value, error?, formula]
    const expected = [
      ['A1', 0, 0, 42, false, undefined],
      ['A2', 1, 0, 45306.5, false, 'DATE(2024,1,15)+0.5'],
      ['A3', 2, 0, '#DIV/0!', true, '1/0'],
      ['A4', 3, 0, 84, false, 'A1*2'],
      ['A5', 4, 0, 'hidden row', false, undefined],
      ['D1', 0, 3, 'Name', false, undefined],
      ['E1', 0, 4, 'Qty', false, undefined],
      ['D2', 1, 3, 'a', false, undefined],
      ['E2', 1, 4, 1, false, undefined],
      ['D3', 2, 3, 'b', false, undefined],
      ['E3', 2, 4, 2, false, undefined],
      ['XFD1048576', 1048575, 16383, 'corner', false, undefined],
    ] as const
    for (const [ref, row, col, value, isError, formula] of expected) {
      const left = await a.readCell({ row, col }, opts(a))
      expect(await b.readCell({ row, col }, opts(b)), `${ref} parity`).toEqual(left)
      const cell = cellOf(left)
      expect(cell.value, ref).toBe(value)
      expect(cell.valueIsError, ref).toBe(isError)
      expect(cell.formula, ref).toBe(formula)
      expect(cell.valueProvenance, ref).toBe('cached')
    }
    // A cell Excel never wrote is blank on both, not unknown.
    const blank = await a.readCell({ row: 6, col: 0 }, opts(a))
    expect(await b.readCell({ row: 6, col: 0 }, opts(b))).toEqual(blank)
    expect(blank.kind).toBe('blank')
    await paged.release()
  })

  it('#DIV/0! is an error token, not error-looking text, in the typed column lanes of both adapters', async () => {
    const { object, paged } = await bothAdapters()
    const query = { col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 5 }] }, limit: 100, encoding: 'lanes' as const }
    const a = object.store('1'), b = paged.store('1')
    const left = await a.queryColumn(query, opts(a)), right = await b.queryColumn(query, opts(b))
    expect({ ...right, stamp: null }).toEqual({ ...left, stamp: null })
    const lanes = left.items[0] as unknown as ColumnLanes
    expect(lanes.kind).toBe('lanes')
    const at = Array.from(lanes.rows).indexOf(2)
    expect(at).toBeGreaterThanOrEqual(0)
    // Lane tag 4 is the error class (laneValue in values.ts); its text is the token.
    expect(lanes.tags[at]).toBe(4)
    expect(String.fromCharCode(...lanes.text.subarray(lanes.textStart[at], lanes.textStart[at + 1]))).toBe('#DIV/0!')
    // A5's string is tag 2 (text), so the error class is not just "any string".
    expect(lanes.tags[Array.from(lanes.rows).indexOf(4)]).toBe(2)
    await paged.release()
  })

  it('A2 keeps serial 45306.5 and formats as Excel displayed it: "2024-01-15 12:00"', async () => {
    const { object, paged, sheet } = await bothAdapters()
    const a = object.store('1'), b = paged.store('1')
    const left = cellOf(await a.readCell({ row: 1, col: 0 }, opts(a)))
    const right = cellOf(await b.readCell({ row: 1, col: 0 }, opts(b)))
    expect(right.styleIndex).toBe(left.styleIndex)
    // Resident display: the parsed sheet and the parsed cell's style.
    const residentCell = sheet.rows.find(r => r.index === 1)?.cells.find(c => c.col === 0)
    expect(residentCell?.style?.formatCode).toBe('yyyy\\-mm\\-dd\\ hh:mm')
    expect(cellDisplayText(sheet, { value: left.value, style: residentCell?.style })).toBe('2024-01-15 12:00')
    // Paged display: the worker's own style table and date system (as viewerFilterEvaluator builds it).
    const formatting = paged.formatting?.()
    expect(formatting?.dateSystem).toBe('1900')
    const workerSheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(formatting!.dateSystem) } as XlsxSheet
    expect(cellDisplayText(workerSheet, { value: right.value, style: formatting!.styles[right.styleIndex] })).toBe('2024-01-15 12:00')
    await paged.release()
  })

  it('A6 publishes Excel\'s exact UTF-16 code units, surrogate pair included, on both adapters', async () => {
    const { object, paged } = await bothAdapters()
    const a = object.store('1'), b = paged.store('1')
    const left = await a.readCell({ row: 5, col: 0 }, opts(a))
    expect(await b.readCell({ row: 5, col: 0 }, opts(b))).toEqual(left)
    const value = cellOf(left).value
    expect(typeof value).toBe('string')
    const text = value as string
    expect(Array.from({ length: text.length }, (_, i) => text.charCodeAt(i))).toEqual(A6_UNITS)
    expect(text).toBe('Ünïcødé 漢字 🎉')
    await paged.release()
  })

  it('row 5 is hidden in base visibility on both adapters; no other row is', async () => {
    const { object, paged } = await bothAdapters()
    const left = object.geometry('1').baseHidden, right = paged.geometry('1').baseHidden
    expect(right.intervals).toEqual(left.intervals)
    expect(left.intervals).toEqual([{ r0: 4, r1: 4 }])
    expect([left.contains(4), right.contains(4)]).toEqual([true, true])
    expect([left.contains(3), left.contains(5), right.contains(3), right.contains(5)]).toEqual([false, false, false, false])
    await paged.release()
  })

  it('the far corner XFD1048576 is inside the used extent and reachable by a row read', async () => {
    const { object, paged } = await bothAdapters()
    const a = object.store('1'), b = paged.store('1')
    expect(b.extent.used).toEqual(a.extent.used)
    expect(a.extent.used).toMatchObject({ r1: 1048575, c1: 16383 })
    const corner = { r0: 1048575, r1: 1048575, c0: 16383, c1: 16383 }
    const left = await a.readRows(corner, undefined, opts(a)), right = await b.readRows(corner, undefined, opts(b))
    expect({ ...right, stamp: null }).toEqual({ ...left, stamp: null })
    expect(JSON.stringify(left.items)).toContain('"corner"')
    await paged.release()
  })

  it('table 表格1 (D1:E3, one header row, no totals) is in the structural index on both adapters', async () => {
    const { object, paged } = await bothAdapters()
    const a = object.store('1'), b = paged.store('1')
    const left = await a.queryStructural(WHOLE, opts(a)), right = await b.queryStructural(WHOLE, opts(b))
    expect(right.tables).toEqual(left.tables)
    expect(left.tables).toHaveLength(1)
    expect(left.tables[0]).toMatchObject({ range: { r0: 0, r1: 2, c0: 3, c1: 4 }, headerRows: 1, totalsRows: 0 })
    for (const result of [left, right]) {
      const index = new StructuralIndex(result)
      const hit = index.firstWhere({ r0: 1, r1: 1, c0: 4, c1: 4 }, h => h.kind === 'table')
      expect(hit?.kind).toBe('table')
    }
    // The resident workbook metadata keeps Excel's zh-TW table name and column names.
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    expect(doc.tables?.map(t => ({ name: t.name, headerRowCount: t.headerRowCount, columns: t.columns.map(c => c.name) })))
      .toEqual([{ name: '表格1', headerRowCount: 1, columns: ['Name', 'Qty'] }])
    await paged.release()
  })

  it('the default engine policy also reproduces Excel\'s displayed values (A2, A3, A4)', async () => {
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const engine = createObjectWorkbookSource(doc, 'engine').store('1')
    const read = async (row: number) => cellOf(await engine.readCell({ row, col: 0 }, opts(engine)))
    expect((await read(1)).value).toBe(45306.5)
    expect(await read(2)).toMatchObject({ value: '#DIV/0!', valueIsError: true })
    expect((await read(3)).value).toBe(84)
  })
})
