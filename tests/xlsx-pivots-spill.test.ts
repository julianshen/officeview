// @vitest-environment node
/**
 * Ruling S3-AO (owner, 2026-11-01), replacing S3-U: a change whose layout would cover a measured blocker — a cell with
 * a value (O1) or another pivot's current area (O3) — commits, and the pivot spills as Excel 16.114 saves it: one
 * #SPILL! cell at its top-left, its other cells empty, `location` that cell, the blocker and the other pivots
 * unchanged, the column widths unchanged, the item state kept. A later change that fits shows the pivot again; Reset
 * shows the saved view. Oracles: interaction/owner-overlap-value.xlsx (O1), owner-overlap-pivot.xlsx (O3).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createResidentPivotClient, pivotAutofitColumns, type PivotCommit } from '../src/components/xlsx/pivot/pivotClient'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import { pivotButtons } from '../src/xlsx/pivots/buttons'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { loadPivotParts } from './helpers/xlsx-pivots-load'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'
import JSZip from 'jszip'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const file = (name: string) => resolve(FIXTURES, `${name}.xlsx`)
const signal = () => new AbortController().signal
const EXPAND: PivotDelta = { kind: 'collapse', field: 0, item: 0, collapsed: false }
const COLLAPSE: PivotDelta = { kind: 'collapse', field: 0, item: 0, collapsed: true }
const SPILL = { value: '#SPILL!', valueIsError: true }
async function load(name: string): Promise<XlsxDocument> { return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file(name))))) }
const pivotSheet = (doc: XlsxDocument) => doc.sheets.findIndex(sheet => sheet.pivotTables?.length)
function putCell(sheet: XlsxSheet, row: number, col: number, value: string) {
  let line = sheet.rows.find(entry => entry.index === row)
  if (!line) { line = { index: row, cells: [] } as unknown as XlsxSheet['rows'][number]; sheet.rows.push(line); sheet.rows.sort((a, b) => a.index - b.index) }
  line.cells.push({ ref: 'A23', row, col, value, styleIndex: 0 } as never)
}
/** layout-collapsed-item (A3:I21; expanding East grows it to A3:I24) with `blocker` in A23: O1's start. */
async function o1(): Promise<{ doc: XlsxDocument; sheet: number }> {
  const doc = await load('engine/layout-collapsed-item'), sheet = pivotSheet(doc)
  putCell(doc.sheets[sheet]!, 22, 0, 'blocker')
  return { doc, sheet }
}
/** The same with owner-overlap-pivot's second pivot at A23:B29 instead: O3's start. */
async function o3(): Promise<{ doc: XlsxDocument; sheet: number; second: number }> {
  const doc = await load('engine/layout-collapsed-item'), sheet = pivotSheet(doc), s = doc.sheets[sheet]!
  const other = (await loadPivotParts(file('interaction/owner-overlap-pivot'))).find(parts => parts.def.name === '樞紐分析表1')!
  s.pivotTables = [...s.pivotTables!, { ...other.def, sheet: s.name }]
  // Its saved cells, as the oracle has them.
  const read = sheetCellReader(other.sheet)
  for (const rect of savedPivotLayout(other.def)!.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
    const cell = read(row, col)
    if (cell) putCell(s, row, col, String(cell.value))
  }
  return { doc, sheet, second: 1 }
}
function resident(doc: XlsxDocument) {
  const installs: PivotCommit[] = []
  return { installs, client: createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => installs.push(commit) }) }
}

describe('S3-AO: an overlapping change spills (resident)', () => {
  it('O1: expanding East into A23 commits as Excel saves it — #SPILL! at A3, the old cells empty, A23 kept, widths kept', async () => {
    const { doc, sheet } = await o1()
    const saved = doc.sheets[sheet]!.pivotTables![0]!
    const target = (await loadPivotParts(file('interaction/owner-overlap-value')))[0]!
    const { client: c, installs } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const spilled = await c.apply(sheet, 0, EXPAND)
    if (!spilled.ok) throw new Error(JSON.stringify(spilled))
    // location = the top-left cell, offsets 0 (Excel's saved <location>).
    expect(spilled.def.location).toMatchObject({ ref: target.def.location.ref, firstHeaderRow: 0, firstDataRow: 0, firstDataCol: 0 })
    expect(spilled.def.location.ref).toBe('A3')
    expect(spilled.def.spill).toEqual({ row: 22, col: 0 })
    // The item state is kept: East expanded, as the oracle's sd.
    expect(spilled.def.fields[0]!.items![0]!.showDetails).toBe(target.def.fields[0]!.items![0]!.showDetails)
    expect(spilled.def.fields[0]!.items![0]!.showDetails).toBe(true)
    const overlay = installs.at(-1)!.overlay!
    expect(overlay.cell(2, 0)).toMatchObject(SPILL)
    // Every other saved cell reads empty; A23 is the sheet's (not the overlay's).
    for (const rect of savedPivotLayout(saved)!.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
      if (row !== 2 || col !== 0) expect(overlay.cell(row, col), `R${row + 1}C${col + 1}`).toBeNull()
    }
    expect(overlay.cell(22, 0)).toBeUndefined()
    expect(sheetCellReader(target.sheet)(22, 0)?.value).toBe('blocker')
    // Accessible notice naming the blocking cell.
    expect(spilled.notice).toMatch(/can't fit: #SPILL!.*A23/)
    // Widths: the oracle's <cols> is the start file's, and the viewer's autofit changes nothing.
    const start = doc.sheets[sheet]!.cols.map(spec => [spec.min, spec.widthChars])
    expect(target.sheet.cols.map(spec => [spec.min, spec.widthChars])).toEqual(start)
    const widths = pivotAutofitColumns({ sheet: doc.sheets[sheet]!, before: { def: saved, read: () => undefined }, after: { def: spilled.def, read: () => undefined }, measure: (_font, text) => text.length * 7, font: () => '11pt x', mdwPx: 7, current: () => undefined, userSet: () => false })
    expect(widths.size).toBe(0)
    // One dropdown on the spilled cell opens the field menu (a change that fits, or Reset).
    expect(pivotButtons(spilled.def, 'screen')).toEqual([{ row: 2, col: 0, kind: 'filter', field: spilled.def.rowFields[0] }])
    expect(pivotButtons(spilled.def, 'print')).toEqual([])
  })

  it('O1: a later change that fits shows the pivot again; Reset shows the saved view', async () => {
    const { doc, sheet } = await o1()
    const saved = doc.sheets[sheet]!.pivotTables![0]!
    const { client: c, installs } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    expect((await c.apply(sheet, 0, EXPAND)).ok).toBe(true)
    const back = await c.apply(sheet, 0, COLLAPSE)
    if (!back.ok) throw new Error(JSON.stringify(back))
    expect(back.def.spill).toBeUndefined()
    expect(back.def.location.ref).toBe('A3:I21')
    expect(installs.at(-1)!.overlay!.cell(2, 0)).not.toMatchObject(SPILL)
    expect((await c.apply(sheet, 0, EXPAND)).ok).toBe(true)
    const reset = await c.apply(sheet, 0, 'reset')
    if (!reset.ok) throw new Error(JSON.stringify(reset))
    expect(reset.def).toBe(saved)
  })

  it('O3: expanding into another pivot spills the first; the second pivot is unchanged', async () => {
    const { doc, sheet, second } = await o3()
    const target = await loadPivotParts(file('interaction/owner-overlap-pivot'))
    const { client: c, installs } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const spilled = await c.apply(sheet, 0, EXPAND)
    if (!spilled.ok) throw new Error(JSON.stringify(spilled))
    expect(spilled.def.location.ref).toBe(target.find(parts => parts.def.name === '樞紐分析表2')!.def.location.ref)
    expect(installs.at(-1)!.overlay!.cell(2, 0)).toMatchObject(SPILL)
    // The second pivot: its definition and its cells are untouched (no overlay over A23:B29).
    expect(doc.sheets[sheet]!.pivotTables![second]!.location.ref).toBe('A23:B29')
    for (let row = 22; row <= 28; row++) for (let col = 0; col <= 1; col++) expect(installs.at(-1)!.overlay!.cell(row, col)).toBeUndefined()
    expect(spilled.notice).toMatch(/A23/)
  })

  it('unmeasured blockers, page fields and linked charts still refuse (fail closed)', async () => {
    // A merge over A23 (S3-AC): refused, naming the first blocked cell.
    const merged = await o1()
    const s = merged.doc.sheets[merged.sheet]!
    s.merges.push('A23:B23'); s.mergeRanges.push({ minRow: 22, maxRow: 22, minCol: 0, maxCol: 1 } as never)
    const { client: c } = resident(merged.doc)
    expect((await c.open(merged.sheet, 0, signal())).ok).toBe(true)
    expect(await c.apply(merged.sheet, 0, EXPAND)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-overlap', message: 'This change would overlap cell A23 on Pivot.' } })
  })
})

let serial = 0
const id = () => `spill-${++serial}`
async function workerRun(source: { kind: 'object'; doc: XlsxDocument } | { kind: 'paged'; bytes: Uint8Array }) {
  const docId = id()
  let kernel: ViewerKernel, sheetId: string
  if (source.kind === 'object') {
    kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
    const sheet = source.doc.sheets.findIndex(candidate => candidate.pivotTables?.length)
    sheetId = source.doc.sheets[sheet]!.sheetId ?? String(sheet + 1)
  } else {
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([source.bytes as BlobPart]) }, { docId, profile: 'D' })
    await imported.done
    const paged = createPagedWorkbookSource(imported)
    kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => paged }, { yield: async () => {} })
    sheetId = paged.sheets.find(sheet => sheet.name === 'Pivot')!.sheetId
  }
  const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
  await run({ kind: 'open', requestId: id(), docId, source: source.kind === 'object' ? { kind: 'object', document: source.doc } : { kind: 'paged', handleId: docId }, profile: 'D' })
  expect(await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })).toMatchObject({ kind: 'complete' })
  let revision = 0, base = 0
  const change = async (delta: PivotDelta | 'reset') => {
    const applied = await run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: base, revision: ++revision, delta, grantBytes: 16 * 1024 * 1024 })
    if (applied.kind !== 'complete' || applied.result.kind !== 'pivotApply') throw new Error(JSON.stringify(applied))
    const committed = await run({ kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: applied.result.revision })
    expect(committed).toMatchObject({ kind: 'complete', result: { kind: 'pivotCommit' } })
    base = applied.result.revision
    return applied.result
  }
  const store = () => kernel.session(docId)!.source.store(sheetId)
  const cell = async (row: number, col: number) => {
    const read = await store().readCell({ row, col }, { signal: signal(), stamp: store().stamp })
    return read.kind === 'cell' ? read.cell : undefined
  }
  return { change, cell, close: () => run({ kind: 'close', requestId: id(), docId }) }
}

describe('S3-AO on the worker (object and paged sources)', () => {
  it('object source: O1 spills, a change that fits shows the pivot again, Reset the saved view', async () => {
    const { doc } = await o1()
    const w = await workerRun({ kind: 'object', doc })
    const spilled = await w.change(EXPAND)
    expect(spilled.def.location.ref).toBe('A3')
    expect(spilled.notice).toMatch(/#SPILL!.*A23/)
    expect(await w.cell(2, 0)).toMatchObject({ value: '#SPILL!' })
    expect((await w.cell(4, 0))?.value ?? null).toBeNull()
    expect(await w.cell(22, 0)).toMatchObject({ value: 'blocker' })
    expect((await w.change(COLLAPSE)).def.location.ref).toBe('A3:I21')
    expect((await w.cell(2, 0))?.value).not.toBe('#SPILL!')
    await w.change(EXPAND)
    expect((await w.change('reset')).def.location.ref).toBe('A3:I21')
    await w.close()
  })

  it('paged source: O1 from the file spills the same way', async () => {
    const zip = await JSZip.loadAsync(readFileSync(file('engine/layout-collapsed-item')))
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file('xl/worksheets/sheet1.xml', xml.replace('</sheetData>', '<row r="23"><c r="A23" t="inlineStr"><is><t>blocker</t></is></c></row></sheetData>'))
    const w = await workerRun({ kind: 'paged', bytes: new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' })) })
    const spilled = await w.change(EXPAND)
    expect(spilled.def.location.ref).toBe('A3')
    expect(await w.cell(2, 0)).toMatchObject({ value: '#SPILL!' })
    expect(await w.cell(22, 0)).toMatchObject({ value: 'blocker' })
    expect((await w.change(COLLAPSE)).def.location.ref).toBe('A3:I21')
    await w.close()
  })
})
