/**
 * S13 native Excel checkpoint, a cap-crossing workbook (AC-STORE-01 and
 * AC-REG-01 native legs, paged path without any override).
 *
 * `native-large.xlsx` was written by Microsoft Excel 16 for macOS (zh-TW UI),
 * driven through AppleScript on the user's machine with their approval
 * (2026-10-13), and reopened in Excel without a repair prompt. A1:C50000 hold
 * =ROW()*10+COLUMN() (Excel stores them as shared formulas), D1:D50000
 * =DATE(2024,1,1)+ROW()+0.5 formatted "yyyy-mm-dd hh:mm", and E50000 "last":
 * 200,001 populated cells, above the 100,000-cell object cap of profile D.
 * The expectations are the values Excel displayed (AppleScript `string value`)
 * or follow from its formulas; a failure is a viewer/Excel disagreement.
 *
 * Privacy edit: the only change from Excel's bytes is the removal of Excel's
 * optional `<mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath …/>`
 * block from xl/workbook.xml; every other zip entry keeps its raw bytes and
 * order. Excel's original native-large.xlsx: 1,800,677 bytes, SHA-256
 * 61a30b0bf2227f94a71849a625c7690335d4ce5ba9184dbc937d339414af8852.
 */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { beforeAll, afterAll, describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { blobRangeSource } from '../src/core/zip-stream'
import { parseXlsx } from '../src/xlsx/parse'
import { cellDisplayText } from '../src/xlsx/render'
import type { XlsxSheet } from '../src/xlsx/types'
import { admitSpreadsheet } from '../src/xlsx/storage/admission'
import { importLimitsFor } from '../src/xlsx/storage/importSession'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import type { WorkbookSource } from '../src/worker/viewerSession'
import type { CellRead, SourceStore } from '../src/xlsx/view/contracts'
import { hasFormula } from '../src/xlsx/view/values'

const BYTES = new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'xlsx-viewer', 'native-large.xlsx')))
const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
function cellOf(read: CellRead) {
  expect(read.kind, JSON.stringify(read)).toBe('cell')
  if (read.kind !== 'cell') throw new Error('unreachable')
  return read.cell
}
const rc = (ref: string) => {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(digits) - 1, col: letters.charCodeAt(0) - 65 }
}
/** D-column serial for 1-based row r: DATE(2024,1,1) is serial 45292 in the 1900 system. */
const dSerial = (r: number) => 45292 + r + 0.5

describe('native Excel checkpoint: native-large.xlsx (200,001 cells, crosses the object cap)', () => {
  let paged: WorkbookSource
  beforeAll(async () => {
    const store = new SourcePageStore({ docId: 'nlarge', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([BYTES as BlobPart]) }, { docId: 'nlarge', profile: 'D', store })
    expect((await session.done).kind).toBe('complete')
    paged = createPagedWorkbookSource(session)
  }, 120_000)
  afterAll(async () => { await paged?.release() })

  it('the fixture bytes are pinned (Excel\'s package minus the absPath element)', async () => {
    expect(BYTES.byteLength).toBe(1_800_516)
    expect(createHash('sha256').update(BYTES).digest('hex')).toBe('32122deb5454d7666bed63867d15395094336b398d7c7b7ef027d8a74e5f7857')
    expect(await (await OfficePackage.load(BYTES)).text('xl/workbook.xml')).not.toContain('absPath')
  })

  it('admission with the shipped profile D limits (no override) picks paged', async () => {
    const admission = await admitSpreadsheet(blobRangeSource(new Blob([BYTES as BlobPart])), 'D', importLimitsFor('D'), new AbortController().signal)
    expect(admission.mode).toBe('paged')
    // The first cap it crosses is the declared sheet-part size (gate 1, from the
    // central directory alone), before any cell is counted.
    if (admission.mode === 'paged') expect(admission.reason).toBe('part xl/worksheets/sheet1.xml declares 13179894 bytes (> 4194304)')
  })

  it('the paged adapter reads Excel\'s values at the named and sampled cells', async () => {
    const store = paged.store('1')
    const expected: Array<[string, string | number]> = [
      ['A1', 11], ['B1', 12], ['C1', 13], ['A50000', 500001], ['B50000', 500002], ['C50000', 500003], ['E50000', 'last'],
      ['A2', 21], ['B12345', 123452], ['C25000', 250003], ['A33333', 333331], ['C49999', 499993],
      ['D1', dSerial(1)], ['D2', dSerial(2)], ['D25000', dSerial(25000)], ['D50000', dSerial(50000)],
    ]
    for (const [ref, value] of expected) expect(cellOf(await store.readCell(rc(ref), opts(store))).value, ref).toBe(value)
    // Shared-formula children are formula cells with a cached value (never
    // "unavailable"). Under the cached policy their formula text is not
    // expanded: they carry only the shared group (si); the master holds the text.
    const child = cellOf(await store.readCell(rc('B12345'), opts(store)))
    expect(child).toMatchObject({ valueProvenance: 'cached', valueIsError: false, formulaType: 'shared' })
    expect(hasFormula(child)).toBe(true)
    expect(cellOf(await store.readCell(rc('A1'), opts(store)))).toMatchObject({ formula: 'ROW()*10+COLUMN()', sharedFormula: { si: 0, ref: 'A1:C20' } })
    // E1 was never written: blank, not unknown.
    expect((await store.readCell(rc('E1'), opts(store))).kind).toBe('blank')
    expect(store.extent.used).toMatchObject({ r0: 0, r1: 49999, c0: 0, c1: 4 })
  })

  it('D2 and D50000 display as Excel showed them (worker formatting path)', async () => {
    const store = paged.store('1')
    const formatting = paged.formatting!()
    expect(formatting.dateSystem).toBe('1900')
    const sheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(formatting.dateSystem) } as XlsxSheet
    const shown = async (ref: string) => {
      const cell = cellOf(await store.readCell(rc(ref), opts(store)))
      return cellDisplayText(sheet, { value: cell.value, style: formatting.styles[cell.styleIndex] })
    }
    expect([await shown('D2'), await shown('D50000')]).toEqual(['2024-01-03 12:00', '2160-11-23 12:00'])
  })

  it('the object adapter (compatibility parse) agrees with the paged adapter at the sampled cells', async () => {
    const object = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(BYTES), { valuePolicy: 'cached' }), 'nlarge-o').store('1')
    const store = paged.store('1')
    for (const ref of ['A1', 'C1', 'D2', 'B12345', 'D25000', 'A50000', 'D50000', 'E50000', 'E1']) {
      expect(await store.readCell(rc(ref), opts(store)), ref).toEqual(await object.readCell(rc(ref), opts(object)))
    }
  }, 120_000)
})
