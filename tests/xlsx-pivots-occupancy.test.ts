// @vitest-environment node
/**
 * Plan Task 29, fix round 1 (Ruling S3-AC): what counts as occupied when a pivot grows. Another pivot's current area
 * always counts, and so does the saved area it vacated while it protects it; so do merges, spill ranges and cells
 * with a comment or note (unmeasured: fail closed). Both paths: the resident client and the worker session.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createResidentPivotClient } from '../src/components/xlsx/pivot/pivotClient'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import type { PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { noteCells } from '../src/xlsx/pivots/stub'
import JSZip from 'jszip'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const file = (name: string) => resolve(FIXTURES, `${name}.xlsx`)
const signal = () => new AbortController().signal
const a1 = (row: number, col: number) => { let name = '', c = col + 1; while (c > 0) { name = String.fromCharCode(65 + (c - 1) % 26) + name; c = Math.floor((c - 1) / 26) } return `${name}${row + 1}` }
const shiftRef = (ref: string, rows: number, cols: number) => ref.split(':').map(part => {
  const m = /^([A-Z]+)(\d+)$/.exec(part.replace(/\$/g, ''))!
  let col = 0
  for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64
  return a1(Number(m[2]) - 1 + rows, col - 1 + cols)
}).join(':')

/**
 * layout-collapsed-item-above-values twice on its sheet, both on the one cache: pivot 0 moved right by 4 columns
 * (E3:I51, Fruit "Apples" collapsed: expanding it grows the pivot down) and pivot 1 moved down by 50 rows (A53:E101;
 * hiding a column item narrows it, vacating E53:E101 — under pivot 0's columns, below its bottom).
 */
async function twoPivots(lower = true): Promise<{ doc: XlsxDocument; sheet: number }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file('engine/layout-collapsed-item-above-values')))))
  const sheet = doc.sheets.findIndex(candidate => candidate.pivotTables?.length), s = doc.sheets[sheet]!
  const def = s.pivotTables![0]!
  const rects = savedPivotLayout(def)!.rects
  const inside = (row: number, col: number) => rects.some(r => row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol)
  const cells = s.rows.flatMap(row => row.cells)
  const out = cells.filter(cell => !inside(cell.row, cell.col))
  const moved = (rows: number, cols: number) => cells.filter(cell => inside(cell.row, cell.col))
    .map(cell => ({ ...cell, row: cell.row + rows, col: cell.col + cols, ref: a1(cell.row + rows, cell.col + cols) }))
  out.push(...moved(0, 4), ...(lower ? moved(50, 0) : []))
  const byRow = new Map<number, typeof cells>()
  for (const cell of out) { let line = byRow.get(cell.row); if (!line) byRow.set(cell.row, line = []); line.push(cell) }
  const template = s.rows[0]!
  s.rows = [...byRow].sort((a, b) => a[0] - b[0]).map(([index, line]) => ({ ...(s.rows.find(row => row.index === index) ?? template), index, cells: line.sort((a, b) => a.col - b.col) }))
  const at = (rows: number, cols: number, name: string): PivotTableDefinition => ({ ...def, name, location: { ...def.location, ref: shiftRef(def.location.ref, rows, cols) } })
  s.pivotTables = lower ? [at(0, 4, 'Upper'), at(50, 0, 'Lower')] : [at(0, 4, 'Upper')]
  s.pivotRegions = s.pivotTables.map(table => ({ ...s.pivotRegions![0]!, ref: table.location.ref }))
  return { doc, sheet }
}
const EXPAND: PivotDelta = { kind: 'collapse', field: 0, item: 0, collapsed: false }
const NARROW: PivotDelta = { kind: 'items', field: 2, hidden: [0] }
function resident(doc: XlsxDocument) {
  const installs: any[] = []
  return { installs, client: createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => installs.push(commit) }) }
}

describe('S3-AC: another pivot\'s vacated saved area', () => {
  it('fixture: the two pivots open; the lower narrows into A53:D101; the upper alone grows into row 52 and beyond', async () => {
    const { doc, sheet } = await twoPivots()
    expect(doc.sheets[sheet]!.pivotTables!.map(def => def.location.ref)).toEqual(['E3:I51', 'A53:E101'])
    const { client: c } = resident(doc)
    expect((await c.open(sheet, 1, signal())).ok).toBe(true)
    const narrowed = await c.apply(sheet, 1, NARROW)
    if (!narrowed.ok) throw new Error(JSON.stringify(narrowed))
    expect(narrowed.def.location.ref).toBe('A53:D101')
  })

  it('resident: the upper pivot cannot grow into the cells the lower one vacated; the lower one resets into them', async () => {
    const { doc, sheet } = await twoPivots()
    const { client: c } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    expect((await c.open(sheet, 1, signal())).ok).toBe(true)
    expect((await c.apply(sheet, 1, NARROW)).ok).toBe(true)
    const refused = await c.apply(sheet, 0, EXPAND)
    expect(refused).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-overlap', message: 'This change would overlap cell E53 on Pivot.' } })
    const reset = await c.apply(sheet, 1, 'reset')
    if (!reset.ok) throw new Error(JSON.stringify(reset))
    expect(reset.def.location.ref).toBe('A53:E101')
    expect(reset.commit.overlay?.cell(52, 4) ?? null).toBeNull()
  })

  it('worker: the same sequence refuses at the vacated E53; the lower pivot then resets into it', async () => {
    const { doc, sheet } = await twoPivots()
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
    let serial = 0
    const id = () => `occupancy-${++serial}`
    const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
    const docId = id(), sheetId = doc.sheets[sheet]!.sheetId ?? String(sheet + 1)
    await run({ kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    let revision = 0
    const change = async (pivot: number, delta: PivotDelta | 'reset', base: number) => {
      const applied = await run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot, baseRevision: base, revision: ++revision, delta, grantBytes: 16 * 1024 * 1024 })
      if (applied.kind !== 'complete' || applied.result.kind !== 'pivotApply') return applied
      return run({ kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot, revision: applied.result.revision })
    }
    for (const pivot of [0, 1]) expect(await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot, grantBytes: 16 * 1024 * 1024 })).toMatchObject({ kind: 'complete' })
    expect(await change(1, NARROW, 0)).toMatchObject({ kind: 'complete', result: { kind: 'pivotCommit' } })
    const refused = await change(0, EXPAND, 0)
    expect(refused.kind).toBe('error')
    expect(JSON.stringify(refused)).toContain('This change would overlap cell E53 on Pivot.')
    expect(await change(1, 'reset', 0)).toMatchObject({ kind: 'complete', result: { kind: 'pivotCommit' } })
    await run({ kind: 'close', requestId: id(), docId })
  })
})

/** Puts a value cell on the sheet (with an array formula's saved range, when given). */
function putCell(sheet: XlsxSheet, row: number, col: number, value: string | number, arrayRef?: string) {
  let line = sheet.rows.find(entry => entry.index === row)
  if (!line) { line = { index: row, cells: [] } as unknown as XlsxSheet['rows'][number]; sheet.rows.push(line); sheet.rows.sort((a, b) => a.index - b.index) }
  line.cells.push({ ref: a1(row, col), row, col, value, styleIndex: 0, ...(arrayRef ? { formula: '1', arrayRef } : {}) } as never)
  line.cells.sort((a, b) => a.col - b.col)
}
const merge = (sheet: XlsxSheet, ref: string) => {
  const [first, last = first] = ref.split(':')
  const at = (part: string) => { const m = /^([A-Z]+)(\d+)$/.exec(part)!; let col = 0; for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64; return { row: Number(m[2]) - 1, col: col - 1 } }
  const a = at(first!), b = at(last)
  sheet.merges.push(ref); sheet.mergeRanges.push({ minRow: a.row, maxRow: b.row, minCol: a.col, maxCol: b.col } as never)
}
/** The upper pivot alone (E3:I51); expanding Apples grows it into row 52 and below, columns E–I. */
const STRUCTURES: Array<[string, (sheet: XlsxSheet) => void, string]> = [
  ['a merge of blank cells (D52:E52, its anchor outside the pivot)', sheet => merge(sheet, 'D52:E52'), 'E52'],
  ['a merge reaching in from below-right (G55:K56)', sheet => merge(sheet, 'G55:K56'), 'G55'],
  ['a spill range\'s blank members (A52 spilling to A52:F52)', sheet => putCell(sheet, 51, 0, 1, 'A52:F52'), 'E52'],
  ['a cell with a comment or note (F52)', sheet => { sheet.noteCells = [51 * 16_384 + 5] }, 'F52'],
]
async function structuredDoc(add: (sheet: XlsxSheet) => void) {
  const { doc, sheet } = await twoPivots(false)
  add(doc.sheets[sheet]!)
  return { doc, sheet }
}
async function workerApply(doc: XlsxDocument, sheet: number): Promise<WorkerReply> {
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  let serial = 0
  const id = () => `structure-${++serial}`
  const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
  const docId = id(), sheetId = doc.sheets[sheet]!.sheetId ?? String(sheet + 1)
  await run({ kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
  expect(await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })).toMatchObject({ kind: 'complete' })
  const reply = await run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: EXPAND, grantBytes: 16 * 1024 * 1024 })
  await run({ kind: 'close', requestId: id(), docId })
  return reply
}

describe('S3-AC: merges, spill ranges and notes count as occupied (fail closed)', () => {
  it('without them the upper pivot grows into rows 52 and below', async () => {
    const { doc, sheet } = await structuredDoc(() => {})
    const { client: c } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    const grown = await c.apply(sheet, 0, EXPAND)
    if (!grown.ok) throw new Error(JSON.stringify(grown))
    const reply = await workerApply((await structuredDoc(() => {})).doc, sheet)
    expect(reply).toMatchObject({ kind: 'complete', result: { kind: 'pivotApply' } })
    if (reply.kind === 'complete' && reply.result.kind === 'pivotApply') expect(reply.result.def.location.ref).toBe(grown.def.location.ref)
    expect(Number(/:[A-Z]+(\d+)$/.exec(grown.def.location.ref)![1])).toBeGreaterThan(56)
  })
  for (const [name, add, cell] of STRUCTURES) {
    it(`resident and worker: ${name} refuses, naming ${cell}`, async () => {
      const message = `This change would overlap cell ${cell} on Pivot.`
      const { doc, sheet } = await structuredDoc(add)
      const { client: c, installs } = resident(doc)
      expect((await c.open(sheet, 0, signal())).ok).toBe(true)
      expect(await c.apply(sheet, 0, EXPAND)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-overlap', message } })
      expect(installs).toHaveLength(0)
      const reply = await workerApply((await structuredDoc(add)).doc, sheet)
      expect(reply.kind).toBe('error')
      expect(JSON.stringify(reply)).toContain(message)
    })
  }
  it('notes that could not be read count everywhere outside the pivot', async () => {
    const { doc, sheet } = await structuredDoc(target => { target.notesUnknown = true })
    const { client: c } = resident(doc)
    expect((await c.open(sheet, 0, signal())).ok).toBe(true)
    expect(await c.apply(sheet, 0, EXPAND)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-overlap', message: 'This change would overlap cell E52 on Pivot.' } })
  })
})

/** layout-collapsed-item (expanding Apples grows A3:I21 to A3:I24) with merges after its sheetData, or a comments part. */
async function injected(options: { merges?: string; comments?: string; threaded?: boolean; array?: string }): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(readFileSync(file('engine/layout-collapsed-item')))
  let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  if (options.array) xml = xml.replace('<sheetData>', `<sheetData><row r="1"><c r="A1"><f t="array" ref="${options.array}">1</f><v>1</v></c></row>`)
  if (options.merges) xml = xml.replace('</sheetData>', `</sheetData><mergeCells count="1"><mergeCell ref="${options.merges}"/></mergeCells>`)
  zip.file('xl/worksheets/sheet1.xml', xml)
  if (options.comments !== undefined) {
    const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')
    const type = options.threaded ? 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels.replace('</Relationships>', `<Relationship Id="rId9" Type="${type}" Target="../comments1.xml"/></Relationships>`))
    zip.file('xl/comments1.xml', options.comments)
  }
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }))
}
const COMMENTS = (ref: string) => `<?xml version="1.0" encoding="UTF-8"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList><comment ref="${ref}" authorId="0"><text><t>note</t></text></comment></commentList></comments>`
const THREADED = (ref: string) => `<?xml version="1.0" encoding="UTF-8"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="${ref}" dT="2026-01-01T00:00:00.00" personId="{0}" id="{1}"><text>hi</text></threadedComment></ThreadedComments>`
async function residentApply(bytes: Uint8Array) {
  const doc = await parseXlsx(await OfficePackage.load(bytes)), sheet = doc.sheets.findIndex(candidate => candidate.pivotTables?.length)
  const { client: c } = resident(doc)
  expect((await c.open(sheet, 0, signal())).ok).toBe(true)
  return { doc, sheet, applied: await c.apply(sheet, 0, EXPAND) }
}
async function pagedApply(bytes: Uint8Array): Promise<WorkerReply> {
  let serial = 0
  const id = () => `paged-structure-${++serial}`, docId = id()
  const imported = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId, profile: 'D' })
  await imported.done
  const source = createPagedWorkbookSource(imported)
  const kernel = new ViewerKernel({ parse: async data => parseXlsx(await OfficePackage.load(data)), openSource: async () => source }, { yield: async () => {} })
  const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
  await run({ kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
  const sheetId = source.sheets.find(sheet => sheet.name === 'Pivot')!.sheetId
  expect(await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })).toMatchObject({ kind: 'complete' })
  const reply = await run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: EXPAND, grantBytes: 16 * 1024 * 1024 })
  await run({ kind: 'close', requestId: id(), docId })
  return reply
}

describe('S3-AC from the file, resident parse and paged import', () => {
  it('noteCells: a comments part\'s cells; a part cut short or naming no cell is unknown', () => {
    expect(noteCells(COMMENTS('B23'))).toEqual([22 * 16_384 + 1])
    expect(noteCells(THREADED('C24'))).toEqual([23 * 16_384 + 2])
    expect(noteCells(COMMENTS('B23').slice(0, -20))).toBeUndefined()
    expect(noteCells(COMMENTS('nowhere'))).toBeUndefined()
    expect(noteCells(undefined)).toBeUndefined()
  })
  const cases: Array<[string, Parameters<typeof injected>[0], string]> = [
    ['a merge of blank cells below the pivot (B23:C23)', { merges: 'B23:C23' }, 'B23'],
    ['an array formula\'s blank cells below the pivot (A1 over A1:A30)', { array: 'A1:A30' }, 'A22'],
    ['a note on a blank cell (C23)', { comments: COMMENTS('C23') }, 'C23'],
    ['a threaded comment on a blank cell (D24)', { comments: THREADED('D24'), threaded: true }, 'D24'],
    ['a comments part that cannot be read (no closing tag): every cell counts', { comments: COMMENTS('Z99').slice(0, -20) }, 'A22'],
  ]
  for (const [name, options, cell] of cases) {
    it(`${name} refuses, naming ${cell}`, async () => {
      const message = `This change would overlap cell ${cell} on Pivot.`
      const bytes = await injected(options)
      const { applied } = await residentApply(bytes)
      expect(applied).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-overlap', message } })
      const reply = await pagedApply(bytes)
      expect(reply.kind).toBe('error')
      expect(JSON.stringify(reply)).toContain(message)
    })
  }
  it('a note far from the pivot changes nothing', async () => {
    const bytes = await injected({ comments: COMMENTS('Z99') })
    const { applied } = await residentApply(bytes)
    if (!applied.ok) throw new Error(JSON.stringify(applied))
    expect(applied.def.location.ref).toBe('A3:I24')
    expect(await pagedApply(bytes)).toMatchObject({ kind: 'complete', result: { kind: 'pivotApply', def: { location: { ref: 'A3:I24' } } } })
  })
})
