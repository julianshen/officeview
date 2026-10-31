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
async function twoPivots(): Promise<{ doc: XlsxDocument; sheet: number }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file('engine/layout-collapsed-item-above-values')))))
  const sheet = doc.sheets.findIndex(candidate => candidate.pivotTables?.length), s = doc.sheets[sheet]!
  const def = s.pivotTables![0]!
  const rects = savedPivotLayout(def)!.rects
  const inside = (row: number, col: number) => rects.some(r => row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol)
  const cells = s.rows.flatMap(row => row.cells)
  const out = cells.filter(cell => !inside(cell.row, cell.col))
  const moved = (rows: number, cols: number) => cells.filter(cell => inside(cell.row, cell.col))
    .map(cell => ({ ...cell, row: cell.row + rows, col: cell.col + cols, ref: a1(cell.row + rows, cell.col + cols) }))
  out.push(...moved(0, 4), ...moved(50, 0))
  const byRow = new Map<number, typeof cells>()
  for (const cell of out) { let line = byRow.get(cell.row); if (!line) byRow.set(cell.row, line = []); line.push(cell) }
  const template = s.rows[0]!
  s.rows = [...byRow].sort((a, b) => a[0] - b[0]).map(([index, line]) => ({ ...(s.rows.find(row => row.index === index) ?? template), index, cells: line.sort((a, b) => a.col - b.col) }))
  const at = (rows: number, cols: number, name: string): PivotTableDefinition => ({ ...def, name, location: { ...def.location, ref: shiftRef(def.location.ref, rows, cols) } })
  s.pivotTables = [at(0, 4, 'Upper'), at(50, 0, 'Lower')]
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

  it('worker: the same sequence refuses, and the lower pivot\'s current area always counts', async () => {
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

export type { XlsxSheet }
