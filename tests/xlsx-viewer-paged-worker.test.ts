/**
 * S1.2 worker wiring: a blob that admission routes to paged opens a paged
 * WorkbookSource (cached policy, no whole-blob copy, no legacy parse) instead
 * of the temporary object source; sort/value queries run over pages; a file
 * the paged path cannot serve yet opens through the compatibility path with a
 * notice; storage refusals are typed open errors.
 */
import { describe, expect, it, vi } from 'vitest'
import type { CacheIdentity, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { fakeIndexedDb } from './helpers/fake-storage'
import { colName, repeat, xlsxBytes } from './helpers/xlsx-package'

const quiet = { yield: async () => {} }
let nextId = 0
const rid = () => `pw-${++nextId}`
async function run<K extends WorkerCommand['kind']>(kernel: ViewerKernel, command: Extract<WorkerCommand, { kind: K }>): Promise<WorkerReply<K>> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, out => { replies.push(out) })
  return replies[replies.length - 1] as WorkerReply<K>
}
function completed<T extends WorkerReply>(out: T): Extract<T, { kind: 'complete' }>['result'] {
  if (out.kind !== 'complete') throw new Error(`expected complete, got ${JSON.stringify(out)}`)
  return (out as Extract<T, { kind: 'complete' }>).result
}
const rowsXml = (values: number[]) => values.map((v, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${v}</v></c><c r="B${r + 1}"><v>${r}</v></c></row>`).join('')
/** Forces admission's paged decision on a tiny file: the object path allows one row. */
const PAGED = { object: { rows: 1 } }

describe('resident worker opens paged sources (S1.2)', () => {
  it('routes a paged admission to PagedXlsxSheetAdapter: no blob copy, no legacy parse, cached policy', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
    const blob = new Blob([xlsxBytes({ sheetData: rowsXml([3, 1, 4, 1, 5]) }) as BlobPart])
    const whole = vi.spyOn(blob, 'arrayBuffer')
    const open = completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'blob', blob }, profile: 'D' }))
    expect(open).toEqual({ kind: 'open', valuePolicy: 'cached', sheets: [{ sheetId: '1', name: 'S' }] })
    expect(parse).not.toHaveBeenCalled()
    expect(whole).not.toHaveBeenCalled()
    const owner: CacheIdentity = initialIdentity('doc', '1')
    const values = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, limit: 10 } }))
    expect(values.kind === 'valueQuery' && values.page.items.map(item => item.kind === 'cell' ? item.cell.value : item.kind)).toEqual([3, 1, 4, 1, 5])
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    const sorted = completed(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: { range: { r0: 0, r1: 4, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } }))
    expect(sorted.kind === 'view' && [...sorted.delta.bands[0].sourceAtOffset]).toEqual([1, 3, 0, 2, 4])
    const session = kernel.session('doc')!
    expect(session.source.valuePolicy).toBe('cached')
    completed(await run(kernel, { kind: 'close', requestId: rid(), docId: 'doc' }))
  })

  it('opens a file the paged path cannot serve yet through the compatibility path with a notice', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
    // Rows without numbers: the legacy parser accepts them, the paged reader refuses.
    const bytes = xlsxBytes({ sheetData: '<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>' })
    const replies: WorkerReply[] = []
    await kernel.handle({ kind: 'open', requestId: rid(), docId: 'compat', source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' }, out => { replies.push(out) })
    const open = replies.at(-1)!
    expect(completed(open)).toMatchObject({ kind: 'open', valuePolicy: 'engine' })
    expect(open.diagnostics.some(d => d.code === 'large-file-compatibility' && /compatibility mode/.test(d.message))).toBe(true)
    expect(parse).toHaveBeenCalledTimes(1)
  })

  it('a large paged document without page storage is a typed storage-error open, not a blank document', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    // The memory fallback holds at most 16 bytes, so this document needs real storage.
    const kernel = new ViewerKernel({ parse, importLimits: { ...PAGED, spool: { memoryLimitBytes: 16 } }, pageStorage: {} }, quiet)
    const out = await run(kernel, { kind: 'open', requestId: rid(), docId: 'nostore', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData: rowsXml([1, 2]) }) as BlobPart]) }, profile: 'D' })
    expect(out).toMatchObject({ kind: 'error', code: 'storage-error', retryable: true })
    expect(kernel.session('nostore')).toBeUndefined()
    expect(parse).not.toHaveBeenCalled()
  })

  it('stores pages in IndexedDB when the document exceeds the memory fallback and removes them on close', async () => {
    const idb = fakeIndexedDb()
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: { ...PAGED, spool: { memoryLimitBytes: 16 } }, pageStorage: { indexedDb: idb } }, quiet)
    const values = Array.from({ length: 50 }, (_, i) => (i * 37) % 11)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'idb', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData: rowsXml(values) }) as BlobPart]) }, profile: 'D' }))
    const owner = initialIdentity('idb', '1')
    const page = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 40, r1: 49 }] }, limit: 20 } }))
    expect(page.kind === 'valueQuery' && page.page.items.map(item => item.kind === 'cell' ? item.cell.value : item.kind)).toEqual(values.slice(40))
    expect(idb.stores.size).toBe(1)
    completed(await run(kernel, { kind: 'close', requestId: rid(), docId: 'idb' }))
    expect(idb.stores.size).toBe(0)
    expect(idb.deleted.length).toBe(1)
  })
})
void colName; void repeat

import { openSpreadsheetSource } from '../src/components/largeSource'
import { importLimitsFor } from '../src/xlsx/storage/importSession'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'

describe('internal spreadsheet entry: paged decisions open paged sessions (the TODO(S1.2) flip)', () => {
  const ctx = (limits = PAGED) => ({ signal: new AbortController().signal, profile: 'D' as const, limits: importLimitsFor('D', limits), options: { limits } })

  it('opens a paged session for a paged admission without reading the whole blob', async () => {
    const blob = new Blob([xlsxBytes({ sheetData: rowsXml([3, 1, 4]) }) as BlobPart])
    const whole = vi.spyOn(blob, 'arrayBuffer')
    const opened = await openSpreadsheetSource(blob, ctx())
    expect(opened.kind).toBe('paged')
    if (opened.kind !== 'paged') return
    expect(whole).not.toHaveBeenCalled()
    expect(await opened.session.done).toMatchObject({ kind: 'complete' })
    const sheet = createPagedWorkbookSource(opened.session).store('1')
    expect(await sheet.readCell({ row: 2, col: 0 }, { signal: new AbortController().signal, stamp: { ...sheet.stamp } })).toMatchObject({ kind: 'cell', cell: { value: 4 } })
    await opened.session.release()
  })

  it('keeps the compatibility path, with its notice, for a file the paged reader cannot serve yet', async () => {
    const blob = new Blob([xlsxBytes({ sheetData: '<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>' }) as BlobPart])
    const opened = await openSpreadsheetSource(blob, ctx())
    expect(opened.kind).toBe('document')
    if (opened.kind !== 'document') return
    expect(opened.document.diagnostics?.map(d => d.message)).toContain('Large file — opened in compatibility mode')
    expect(opened.document.sheets[0].rows[0].cells[0].value).toBe(1)
  })

  it('object admissions load the document unchanged', async () => {
    const opened = await openSpreadsheetSource(new Blob([xlsxBytes({ sheetData: rowsXml([7]) }) as BlobPart]), ctx({} as typeof PAGED))
    expect(opened.kind).toBe('document')
  })

  it('a paged admission with no page storage is a typed refusal, never an in-memory fallback', async () => {
    const limits = { ...PAGED, spool: { memoryLimitBytes: 16 } }
    await expect(openSpreadsheetSource(new Blob([xlsxBytes({ sheetData: rowsXml([1, 2]) }) as BlobPart]), ctx(limits), {})).rejects.toMatchObject({ name: 'ImportRefusal', code: 'storage-error' })
  })
})

describe('RED 14a through the worker: unreadable structure refuses sort, never sorts past it', () => {
  it('a paged sheet whose mergeCells sit in mc:AlternateContent refuses a sort over it with the structural reason', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    const bytes = xlsxBytes({ sheetData: rowsXml([3, 1, 4, 1, 5]), afterSheetData: `<mc:AlternateContent ${MC}><mc:Choice Requires="x"><mergeCells><mergeCell ref="A2:B3"/></mergeCells></mc:Choice></mc:AlternateContent>` })
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'mc', source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' }))
    const owner = initialIdentity('mc', '1')
    const out = await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: { range: { r0: 0, r1: 4, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } })
    expect(out).toMatchObject({ kind: 'error', code: 'unsupported', retryable: false })
    expect(out.diagnostics).toEqual([{ code: 'source-error', message: 'merges metadata is wrapped in mc:AlternateContent and cannot be read' }])
    // Inspection is unaffected.
    const values = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, limit: 10 } }))
    expect(values.kind === 'valueQuery' && values.page.items.length).toBe(5)
  })
})

describe('review MINOR 1: the compatibility fallback is bounded by the object caps', () => {
  const NO_ROW_NUMBERS = () => xlsxBytes({ sheetData: '<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>' })
  const BEYOND = { object: { rows: 1, compressedBytes: 100 } }

  it('worker: beyond the object compressed cap, an unservable file stays a typed incomplete paged source (no legacy parse)', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: BEYOND }, quiet)
    const replies: WorkerReply[] = []
    await kernel.handle({ kind: 'open', requestId: rid(), docId: 'big', source: { kind: 'blob', blob: new Blob([NO_ROW_NUMBERS() as BlobPart]) }, profile: 'D' }, out => { replies.push(out) })
    const open = replies.at(-1)!
    expect(completed(open)).toMatchObject({ kind: 'open', valuePolicy: 'cached' })
    expect(open.diagnostics).toContainEqual(expect.objectContaining({ code: 'paged-unsupported' }))
    expect(parse).not.toHaveBeenCalled()
    const owner = initialIdentity('big', '1')
    const out = await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 10 } })
    expect(JSON.stringify(out)).toContain('unsupported')
  })

  it('internal entry: beyond the cap the paged session is returned incomplete instead of a legacy parse', async () => {
    const limits = BEYOND as unknown as typeof PAGED
    const opened = await openSpreadsheetSource(new Blob([NO_ROW_NUMBERS() as BlobPart]), { signal: new AbortController().signal, profile: 'D', limits: importLimitsFor('D', limits), options: { limits } })
    expect(opened.kind).toBe('paged')
    if (opened.kind !== 'paged') return
    expect(opened.session.completeness).toBe('incomplete')
    expect((await opened.session.done)).toMatchObject({ kind: 'incomplete', failure: { code: 'unsupported' } })
    await opened.session.release()
  })
})
