/**
 * S13 native Excel checkpoint, dates (AC-STORE-01 and AC-SEM-07 native legs).
 *
 * `native-1904.xlsx` was written by Microsoft Excel 16 for macOS (zh-TW UI)
 * with the 1904 date system on, driven through AppleScript on the user's
 * machine with their approval (2026-10-13), and reopened in Excel without a
 * repair prompt. The expectations are the values Excel displayed after
 * reopening (AppleScript `string value`). A failure is a viewer/Excel
 * disagreement to report, never an expectation to relax.
 *
 * Privacy edit: the only change from Excel's bytes is the removal of Excel's
 * optional `<mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath …/>`
 * block (the local save folder) from xl/workbook.xml; every other zip entry
 * keeps its raw bytes and order. Excel's original native-1904.xlsx: 9,658
 * bytes, SHA-256 12ccf94d54aec7bb85e781524fc8c9241b668ad727c23edef8a6ffd9c051fc84.
 *
 * The date filters run through the worker filter evaluator (the production
 * rule path for worker-owned sheets: classification, FilterEngine date
 * lowering to [start, nextStart), value lanes) over BOTH adapters' stores.
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
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { createWorkerFilterEvaluator } from '../src/worker/viewerFilterEvaluator'
import type { WorkbookSource } from '../src/worker/viewerSession'
import type { CellRead, ColumnRule, FilterSpec, SourceStore, WorkbookFormatting } from '../src/xlsx/view/contracts'

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'xlsx-viewer', name)))
const BYTES_1904 = fixture('native-1904.xlsx')
const BYTES_1900 = fixture('native-corner.xlsx')

const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })

async function bothAdapters(bytes: Uint8Array, docId: string): Promise<{ object: WorkbookSource; paged: WorkbookSource; sheet: XlsxSheet }> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const object = createObjectWorkbookSource(doc, docId)
  const store = new SourcePageStore({ docId, backend: memoryPageBackend() })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId, profile: 'D', store, pageRows: 4 })
  expect((await session.done).kind).toBe('complete')
  return { object, paged: createPagedWorkbookSource(session), sheet: doc.sheets[0] }
}
function cellOf(read: CellRead) {
  expect(read.kind, JSON.stringify(read)).toBe('cell')
  if (read.kind !== 'cell') throw new Error('unreachable')
  return read.cell
}

/** Column A as Excel saved and displayed it: [row, formula, 1904 serial, Excel's display]. */
const DATES_1904 = [
  [1, 'DATE(2024,1,15)+0.5', 43844.5, '2024-01-15 12:00:00'],
  [2, 'DATE(2024,1,16)', 43845, '2024-01-16 00:00:00'],
  [3, 'DATE(2024,1,16)+0.999988426', 43845.999988426003, '2024-01-16 23:59:59'],
  [4, 'DATE(2024,2,29)+0.25', 43889.25, '2024-02-29 06:00:00'],
  [5, 'DATE(1904,1,1)', 0, '1904-01-01 00:00:00'],
  [6, 'DATE(2023,12,31)+0.5', 43829.5, '2023-12-31 12:00:00'],
] as const

describe('native Excel checkpoint: native-1904.xlsx (date1904 workbook)', () => {
  it('the fixture bytes are pinned (Excel\'s package minus the absPath element)', async () => {
    expect(BYTES_1904.byteLength).toBe(9_499)
    expect(createHash('sha256').update(BYTES_1904).digest('hex')).toBe('eebaea60f40ba35bef3e4941ad862317f71f8abfd68af2109f2d3a34942a81dc')
    const workbook = await (await OfficePackage.load(BYTES_1904)).text('xl/workbook.xml')
    expect(workbook).not.toContain('absPath')
    expect(workbook).toContain('<workbookPr date1904="1"')
  })

  it('both adapters agree with Excel: 1904 serials, formulas, epoch cell A6, header and labels', async () => {
    const { object, paged, sheet } = await bothAdapters(BYTES_1904, 'n1904')
    expect(paged.sheets).toEqual(object.sheets)
    expect(object.sheets[0].name).toBe('工作表1')
    expect(sheet.semantics?.dateSystem).toBe('1904')
    expect(paged.formatting?.().dateSystem).toBe('1904')
    const a = object.store('1'), b = paged.store('1')
    for (const [row, formula, serial] of DATES_1904) {
      const left = await a.readCell({ row, col: 0 }, opts(a))
      expect(await b.readCell({ row, col: 0 }, opts(b)), `A${row + 1} parity`).toEqual(left)
      expect(cellOf(left), `A${row + 1}`).toMatchObject({ value: serial, formula, valueIsError: false, valueProvenance: 'cached' })
    }
    for (const [col, text] of [[0, 'When'], [1, 'Label']] as const) {
      const left = await a.readCell({ row: 0, col }, opts(a))
      expect(await b.readCell({ row: 0, col }, opts(b))).toEqual(left)
      expect(cellOf(left).value).toBe(text)
    }
    await paged.release()
  })

  it('display text matches Excel through the resident and the worker formatting paths', async () => {
    const { object, paged, sheet } = await bothAdapters(BYTES_1904, 'n1904d')
    const a = object.store('1'), b = paged.store('1')
    const formatting = paged.formatting!()
    const workerSheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(formatting.dateSystem) } as XlsxSheet
    const shown: string[][] = []
    for (const [row, , , excel] of DATES_1904) {
      const left = cellOf(await a.readCell({ row, col: 0 }, opts(a))), right = cellOf(await b.readCell({ row, col: 0 }, opts(b)))
      const style = sheet.rows.find(r => r.index === row)?.cells.find(c => c.col === 0)?.style
      shown.push([`A${row + 1}`, excel, cellDisplayText(sheet, { value: left.value, style }), cellDisplayText(workerSheet, { value: right.value, style: formatting.styles[right.styleIndex] })])
    }
    // One table so a disagreement names every cell at once.
    expect(shown).toEqual(shown.map(([ref, excel]) => [ref, excel, excel, excel]))
    await paged.release()
  })
})

describe('AC-SEM-07 native leg: date filters over Excel-written cells', () => {
  // Local noon on 2024-02-15, so thisMonth/thisYear resolve to Feb 2024 / 2024 in any zone.
  const nowMs = new Date(2024, 1, 15, 12).getTime()
  const day = (year: number, month: number, d: number) => ({ year, month, day: d })

  async function matching(source: WorkbookSource, formatting: WorkbookFormatting, range: FilterSpec['range'], rule: ColumnRule): Promise<number[]> {
    const store = source.store('1')
    const spec: FilterSpec = { range, columns: new Map([[0, rule]]) }
    const result = await createWorkerFilterEvaluator(() => new Map(), formatting)({ store, spec, nowMs, signal: new AbortController().signal })
    if (result.kind !== 'complete') throw new Error(`filter ${JSON.stringify(rule)}: ${JSON.stringify(result)}`)
    const rows: number[] = []
    for (let row = range.r0 + range.headerRows; row <= range.r1; row++) if (result.matches(row, spec)) rows.push(row)
    return rows
  }

  it('1904 workbook: eq/before/after/between/thisMonth/thisYear lower to [start, nextStart) on both adapters', async () => {
    const { object, paged } = await bothAdapters(BYTES_1904, 'n1904f')
    const formatting = paged.formatting!()
    const range = { r0: 0, r1: 6, c0: 0, c1: 1, headerRows: 1 as const, totalsRows: 0 as const }
    // Source rows: 1 Jan 15 noon, 2 Jan 16 00:00, 3 Jan 16 23:59:59, 4 Feb 29 06:00, 5 epoch 1904-01-01, 6 Dec 31 2023 noon.
    const cases: Array<[string, ColumnRule, number[]]> = [
      ['equals 2024-01-16 (midnight and 23:59:59, not Jan 15 noon)', { kind: 'date', op: 'eq', first: day(2024, 1, 16) }, [2, 3]],
      ['before 2024-01-16', { kind: 'date', op: 'before', first: day(2024, 1, 16) }, [1, 5, 6]],
      ['after 2024-01-16 (from Jan 17 00:00)', { kind: 'date', op: 'after', first: day(2024, 1, 16) }, [4]],
      ['between 2024-02-01 and the leap day 2024-02-29 (end day inclusive)', { kind: 'date', op: 'between', first: day(2024, 2, 1), second: day(2024, 2, 29) }, [4]],
      ['equals the 1904 epoch 1904-01-01 (serial 0)', { kind: 'date', op: 'eq', first: day(1904, 1, 1) }, [5]],
      ['this month (Feb 2024)', { kind: 'datePreset', preset: 'thisMonth' }, [4]],
      ['this year (2024)', { kind: 'datePreset', preset: 'thisYear' }, [1, 2, 3, 4]],
    ]
    for (const [label, rule, expected] of cases) {
      expect(await matching(object, formatting, range, rule), `object: ${label}`).toEqual(expected)
      expect(await matching(paged, formatting, range, rule), `paged: ${label}`).toEqual(expected)
    }
    await paged.release()
  })

  it('1900 workbook (native-corner.xlsx A2, 2024-01-15 12:00): its noon serial matches its own day only', async () => {
    const { object, paged } = await bothAdapters(BYTES_1900, 'n1900f')
    const formatting = paged.formatting!()
    expect(formatting.dateSystem).toBe('1900')
    // A2 alone: column A also holds text and an error, which would make it a text column.
    const range = { r0: 1, r1: 1, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }
    const cases: Array<[string, ColumnRule, number[]]> = [
      ['equals 2024-01-15', { kind: 'date', op: 'eq', first: day(2024, 1, 15) }, [1]],
      ['equals 2024-01-16', { kind: 'date', op: 'eq', first: day(2024, 1, 16) }, []],
      ['before 2024-01-15', { kind: 'date', op: 'before', first: day(2024, 1, 15) }, []],
      ['between 2024-01-01 and 2024-01-15', { kind: 'date', op: 'between', first: day(2024, 1, 1), second: day(2024, 1, 15) }, [1]],
      ['this year (2024)', { kind: 'datePreset', preset: 'thisYear' }, [1]],
    ]
    for (const [label, rule, expected] of cases) {
      expect(await matching(object, formatting, range, rule), `object: ${label}`).toEqual(expected)
      expect(await matching(paged, formatting, range, rule), `paged: ${label}`).toEqual(expected)
    }
    await paged.release()
  })
})
