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
    expect(open).toMatchObject({ kind: 'open', valuePolicy: 'cached', sheets: [{ sheetId: '1', name: 'S' }], formatting: { dateSystem: '1900' } })
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

import { loadOfficeFile } from '../src/components/OfficeFile'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'

describe('coordinator decision on MINOR 1: the exported entry keeps the earlier rule', () => {
  it('a >32 MiB file the paged reader refuses still opens through compatibility mode (legacy parse) on loadOfficeFile', async () => {
    const entries = xlsxEntries({ sheetData: '<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>' })
    // A stored 33 MiB media part puts the package beyond the 32 MiB object cap (G1 → paged).
    entries.push({ name: 'xl/media/pad.png', data: new Uint8Array(33 * 1024 * 1024), method: 0 })
    const bytes = buildZip(entries).bytes
    expect(bytes.length).toBeGreaterThan(32 * 1024 * 1024)
    const doc = await loadOfficeFile(new Blob([bytes as BlobPart])) as import('../src/xlsx/types').XlsxDocument
    expect(doc.diagnostics?.map(d => d.message)).toContain('Large file — opened in compatibility mode')
    expect(doc.sheets[0].rows.flatMap(r => r.cells.map(c => c.value))).toEqual([1, 2])
  }, 60_000)
})

describe('review round 2 NIT 1: the fallback is also bounded by the declared expanded size', () => {
  it('a small compressed package whose declared expanded size exceeds the object budget does not fall back to the legacy parse', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const limits = { object: { rows: 1, budgetBytes: 1000 } }
    const kernel = new ViewerKernel({ parse, importLimits: limits }, quiet)
    const replies: WorkerReply[] = []
    const bytes = xlsxBytes({ sheetData: '<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>' })
    await kernel.handle({ kind: 'open', requestId: rid(), docId: 'exp', source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' }, out => { replies.push(out) })
    expect(completed(replies.at(-1)!)).toMatchObject({ kind: 'open', valuePolicy: 'cached' })
    expect(replies.at(-1)!.diagnostics).toContainEqual(expect.objectContaining({ code: 'paged-unsupported' }))
    expect(parse).not.toHaveBeenCalled()
  })
})

// ---------------------------------------------------------------------------
// S5.4: bounded worker commands behind the paged interactive viewer
// ---------------------------------------------------------------------------
import { KEY_NUMBER } from '../src/xlsx/view/valueRuns'
import { valueScope } from '../src/xlsx/view/valueIndex'
import type { FilterSpec, SourceReadQuery } from '../src/xlsx/view/contracts'
import { READ_ROWS_PAGE_ROWS } from '../src/xlsx/view/source'

const CORNER_ROWS = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'
  + '<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>corner</t></is></c></row>'
async function openPaged(docId: string, sheetData: string) {
  const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
  const kernel = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
  completed(await run(kernel, { kind: 'open', requestId: rid(), docId, source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData }) as BlobPart]) }, profile: 'D' }))
  await kernel.session(docId)!.source.store('1').ensureComplete({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, { signal: new AbortController().signal, stamp: { docId, sheetId: '1', valueGen: 0 } })
  return { kernel, parse, owner: initialIdentity(docId, '1') }
}
const read = (kernel: ViewerKernel, owner: CacheIdentity, query: SourceReadQuery) => run(kernel, { kind: 'read', requestId: rid(), owner, query })
function readResult(out: WorkerReply<'read'>) {
  const result = completed(out)
  if (result.kind !== 'read') throw new Error(`expected read, got ${result.kind}`)
  return result.result
}

describe('S5.4 bounded worker reads for the paged interactive viewer', () => {
  it('activation publishes the compact sheet layout; cell, row and scan reads reach XFD1048576', async () => {
    const { kernel, parse, owner } = await openPaged('corner', CORNER_ROWS)
    const active = completed(await run(kernel, { kind: 'activateSheet', requestId: rid(), owner: kernel.session('corner')!.workbookOwner(), sheetId: '1' }))
    expect(active).toMatchObject({ kind: 'activateSheet', sheetId: '1', layout: { rowOverrides: [], colOverrides: [], hidden: [], used: { r0: 0, r1: 1048575, c0: 0, c1: 16383 } } })
    // The same layout is re-readable later (it grows while an import runs).
    expect(readResult(await read(kernel, owner, { op: 'layout' }))).toEqual({ op: 'layout', layout: active.kind === 'activateSheet' ? active.layout : undefined })
    const cell = readResult(await read(kernel, owner, { op: 'cell', at: { row: 1048575, col: 16383 } }))
    expect(cell).toMatchObject({ op: 'cell', read: { kind: 'cell', cell: { row: 1048575, col: 16383, value: 'corner' } } })
    const rows = readResult(await read(kernel, owner, { op: 'rows', range: { r0: 1048570, r1: 1048575, c0: 16380, c1: 16383 } }))
    expect(rows.op === 'rows' && rows.page.completeness).toBe('complete')
    expect(rows.op === 'rows' && rows.page.items.map(item => [item.row, item.cells.map(c => c.value)])).toEqual([[1048575, ['corner']]])
    const first = readResult(await read(kernel, owner, { op: 'scan', request: { range: { r0: 0, r1: 1048575, c0: 0, c1: 16383 }, limit: 2 } }))
    expect(first.op === 'scan' && first.page.items.map(c => c.value)).toEqual([1, 2])
    const cursor = first.op === 'scan' ? first.page.nextCursor : undefined
    expect(cursor).toEqual(expect.any(String))
    const rest = readResult(await read(kernel, owner, { op: 'scan', request: { range: { r0: 0, r1: 1048575, c0: 0, c1: 16383 }, limit: 2 }, cursor }))
    expect(rest.op === 'scan' && rest.page.items.map(c => c.value)).toEqual(['corner'])
    expect(rest.op === 'scan' && rest.page.nextCursor).toBeUndefined()
    expect(parse).not.toHaveBeenCalled()
  })

  it('a whole-sheet row read is one bounded page with a continuation, never the full sheet', async () => {
    const values = Array.from({ length: READ_ROWS_PAGE_ROWS + 50 }, (_, i) => i)
    const { kernel, owner } = await openPaged('bounded', rowsXml(values))
    const page = readResult(await read(kernel, owner, { op: 'rows', range: { r0: 0, r1: 1048575, c0: 0, c1: 16383 } }))
    if (page.op !== 'rows') throw new Error('rows')
    expect(page.page.items.length).toBeLessThanOrEqual(READ_ROWS_PAGE_ROWS)
    expect(page.page.nextCursor).toEqual(expect.any(String))
    // An invalid rectangle is a typed failure, not a thrown worker exception.
    const bad = readResult(await read(kernel, owner, { op: 'rows', range: { r0: 5, r1: 1, c0: 0, c1: 0 } }))
    expect(bad.op === 'rows' && bad.page.error?.code).toBe('invalid-source-rect')
  })

  it('reads are owned by the committed view: stale after a commit, cancelled on cancel, withheld after close', async () => {
    const { kernel, owner } = await openPaged('owned', rowsXml([3, 1, 4, 1, 5]))
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    const sorted = completed(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: { range: { r0: 0, r1: 4, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } }))
    const next = sorted.kind === 'view' ? sorted.delta.owner : owner
    expect(await read(kernel, owner, { op: 'cell', at: { row: 0, col: 0 } })).toMatchObject({ kind: 'error', code: 'stale-source' })
    expect(readResult(await read(kernel, next, { op: 'cell', at: { row: 0, col: 0 } }))).toMatchObject({ read: { kind: 'cell', cell: { value: 3 } } })

    const cancelId = rid()
    const replies: WorkerReply[] = []
    const cancelling = new ViewerKernel({ parse: async () => { throw new Error('no parse') }, importLimits: PAGED, afterDispatch: async command => {
      if (command.kind === 'read') await cancelling.handle({ kind: 'cancel', requestId: rid(), targetRequestId: cancelId }, () => {})
    } }, quiet)
    completed(await run(cancelling, { kind: 'open', requestId: rid(), docId: 'c', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData: rowsXml([1, 2]) }) as BlobPart]) }, profile: 'D' }))
    await cancelling.handle({ kind: 'read', requestId: cancelId, owner: initialIdentity('c', '1'), query: { op: 'cell', at: { row: 0, col: 0 } } }, out => { replies.push(out) })
    expect(replies.at(-1)).toMatchObject({ kind: 'cancelled', commandKind: 'read' })

    const closing = new ViewerKernel({ parse: async () => { throw new Error('no parse') }, importLimits: PAGED, afterDispatch: async command => {
      if (command.kind === 'read') await closing.handle({ kind: 'close', requestId: rid(), docId: 'x' }, () => {})
    } }, quiet)
    completed(await run(closing, { kind: 'open', requestId: rid(), docId: 'x', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData: rowsXml([1, 2]) }) as BlobPart]) }, profile: 'D' }))
    const late = await read(closing, initialIdentity('x', '1'), { op: 'rows', range: { r0: 0, r1: 0, c0: 0, c1: 0 } })
    expect(late.kind).not.toBe('complete')
  })

  it('commits typed and value-list filters through the worker on a paged sheet (keys shipped with the command)', async () => {
    const { kernel, owner } = await openPaged('filter', rowsXml([3, 1, 4, 1, 5]))
    const range = { r0: 0, r1: 4, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const typed: FilterSpec = { range, columns: new Map([[0, { kind: 'number', op: 'gt', first: 2 }]]) }
    const out = completed(await run(kernel, { kind: 'filter', requestId: rid(), owner, spec: typed, nowMs: 0 }))
    if (out.kind !== 'view') throw new Error('view')
    expect(out.delta.counts.visible).toBe(1048576 - 2)
    expect(out.delta.visibility).toHaveLength(1)
    expect([out.delta.visibility[0].r0, out.delta.visibility[0].r1, out.delta.visibility[0].words[0]]).toEqual([0, 3, 0b0101])

    const ref = { refId: 'keys-1', count: 1, owner: { docId: 'filter', sheetId: '1', valueGen: 0 } }
    const scope = valueScope({ docId: 'filter', sheetId: '1', valueGen: 0, range, col: 0, classification: { kind: 'number', hasNumeric: true }, classificationGen: 0 })
    const values: FilterSpec = { range, columns: new Map([[0, { kind: 'values', selection: { mode: 'include', keys: ref }, showBlanks: false }]]) }
    const refused = await run(kernel, { kind: 'filter', requestId: rid(), owner: out.delta.owner, spec: values, nowMs: 0 })
    expect(refused.kind).toBe('error')
    const kept = completed(await run(kernel, { kind: 'filter', requestId: rid(), owner: out.delta.owner, spec: values, nowMs: 0,
      keySets: [{ refId: 'keys-1', scope, run: { count: 1, kinds: Uint8Array.of(KEY_NUMBER), numbers: Float64Array.of(1), texts: [undefined] } }] }))
    if (kept.kind !== 'view') throw new Error('view')
    // Only the two 1s stay visible: source rows 0, 2 and 4 are hidden.
    expect([kept.delta.visibility[0].r0, kept.delta.visibility[0].r1, kept.delta.visibility[0].words[0]]).toEqual([0, 4, 0b01010])
  })
})

describe('S5.4 review: worker filter semantics and key lifetime', () => {
  it('F1 date rules classify paged columns through the workbook style table', async () => {
    const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
    const kernel = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
    const styles = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'
    const sheetData = '<row r="1"><c r="A1" s="1"><v>44927</v></c></row><row r="2"><c r="A2" s="1"><v>44928</v></c></row><row r="3"><c r="A3" s="1"><v>44929</v></c></row>'
    const open = completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'dates', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData, styles }) as BlobPart]) }, profile: 'D' }))
    expect(open.kind === 'open' && open.formatting?.styles[1]).toMatchObject({ numFmtId: 14 })
    const owner = initialIdentity('dates', '1')
    const spec: FilterSpec = { range: { r0: 0, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'date', op: 'eq', first: { year: 2023, month: 1, day: 2 } }]]) }
    const out = completed(await run(kernel, { kind: 'filter', requestId: rid(), owner, spec, nowMs: 0 }))
    if (out.kind !== 'view') throw new Error('view')
    // Only 2023-01-02 (serial 44928, source row 1) stays visible.
    expect([out.delta.visibility[0].r0, out.delta.visibility[0].r1, out.delta.visibility[0].words[0]]).toEqual([0, 2, 0b010])
  })

  it('F5 a superseded filter finishing never removes the keys its replacement shares', async () => {
    const { kernel, owner } = await openPaged('shared', rowsXml([3, 1, 4, 1, 5]))
    const range = { r0: 0, r1: 4, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const ref = { refId: 'keys-shared', count: 1, owner: { docId: 'shared', sheetId: '1', valueGen: 0 } }
    const scope = valueScope({ docId: 'shared', sheetId: '1', valueGen: 0, range, col: 0, classification: { kind: 'number', hasNumeric: true }, classificationGen: 0 })
    const keySets = () => [{ refId: 'keys-shared', scope, run: { count: 1, kinds: Uint8Array.of(KEY_NUMBER), numbers: Float64Array.of(1), texts: [undefined] } }]
    const spec = (showBlanks: boolean): FilterSpec => ({ range, columns: new Map([[0, { kind: 'values', selection: { mode: 'include', keys: ref }, showBlanks }]]) })
    const first: WorkerReply[] = [], second: WorkerReply[] = []
    const a = kernel.handle({ kind: 'filter', requestId: rid(), owner, spec: spec(false), nowMs: 0, keySets: keySets() }, out => { first.push(out) })
    const b = kernel.handle({ kind: 'filter', requestId: rid(), owner, spec: spec(true), nowMs: 0, keySets: keySets() }, out => { second.push(out) })
    await Promise.all([a, b])
    expect(first.at(-1)).toMatchObject({ kind: 'cancelled' })
    expect(second.at(-1)).toMatchObject({ kind: 'complete' })
  })
})

describe('S10.2: repeated worker filters do not re-read the column (filter apply ~2.8 s on F2)', () => {
  it('pages are read as typed lanes, and a repeated identical value filter reads the column once, not three times', async () => {
    const rows = Array.from({ length: 9000 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>${['alpha', 'beta', 'gamma'][r % 3]}</t></is></c></row>`).join('')
    const { kernel, owner: start } = await openPaged('repeat', rows)
    const store = kernel.session('repeat')!.source.store('1')
    const requests: Array<{ encoding?: string }> = []
    const proto = Object.getPrototypeOf(store) as { queryColumn: (...args: unknown[]) => unknown }
    const original = proto.queryColumn
    const spy = vi.spyOn(proto, 'queryColumn').mockImplementation(function (this: unknown, ...args: unknown[]) {
      requests.push(args[0] as { encoding?: string })
      return original.apply(this, args)
    })
    try {
      const range = { r0: 0, r1: 8999, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }
      const scope = valueScope({ docId: 'repeat', sheetId: '1', valueGen: 0, range, col: 0, classification: { kind: 'text', hasNumeric: false }, classificationGen: 0 })
      const apply = async (owner: CacheIdentity, n: number) => {
        const ref = { refId: `keys-${n}`, count: 1, owner: { docId: 'repeat', sheetId: '1', valueGen: 0 } }
        const spec: FilterSpec = { range, columns: new Map([[0, { kind: 'values', selection: { mode: 'include', keys: ref }, showBlanks: false }]]) }
        const out = completed(await run(kernel, { kind: 'filter', requestId: rid(), owner, spec, nowMs: 0, keySets: [{ refId: `keys-${n}`, scope, run: { count: 1, kinds: Uint8Array.of(2), numbers: Float64Array.of(0), texts: ['alpha'] } }] }))
        if (out.kind !== 'view') throw new Error('view')
        return out.delta
      }
      const before = requests.length
      const first = await apply(start, 1)
      const firstReads = requests.length - before
      const cleared = completed(await run(kernel, { kind: 'filter', requestId: rid(), owner: first.owner, spec: null, nowMs: 0 }))
      if (cleared.kind !== 'view') throw new Error('view')
      const mark = requests.length
      const second = await apply(cleared.delta.owner, 2)
      const secondReads = requests.length - mark
      expect(second.counts.visible).toBe(first.counts.visible)
      // 9,000 rows = 3 lane pages per column pass. The first apply makes three
      // passes (unavailable-value guard, classification, rule evaluation); the
      // repeat reuses the guard's and the classification's outcome.
      expect(firstReads).toBe(9)
      expect(secondReads).toBe(3)
      // Every page crosses as typed lanes, never as one object per cell.
      expect(requests.filter(request => request.encoding !== 'lanes')).toEqual([])
    } finally { spy.mockRestore() }
  })
})
