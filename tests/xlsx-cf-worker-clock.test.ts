// @vitest-environment node
/**
 * Stage 5 final review F3: the worker (paged) CF preparation samples the clock lazily, as the resident evaluation context
 * does (Task 39 fix round 1, F5). A sheet whose rules never read the clock (a literal `cellIs`) takes no clock sample; a
 * sheet whose rules do takes exactly one per preparation, shared by the predicates, the formulas and the stamp (S5-M).
 */
import { describe, expect, it } from 'vitest'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { GREEN, NOW, cellIsRule, cfBlock, exprRule, rcOf, withPaint, workbookBytes } from './helpers/cf-views'

const CELLS = { A2: 1, A3: 2, A4: 3, B2: 1, B3: 5, B4: 9 }
const timePeriodRule = (priority: number, dxfId: number): string =>
  `<cfRule type="timePeriod" priority="${priority}" dxfId="${dxfId}" timePeriod="today"><formula>FLOOR(A2,1)=TODAY()</formula></cfRule>`

async function workerRead(bytes: Uint8Array, refs: readonly string[]): Promise<{ clockReads: number; fills: (string | undefined)[]; kind: string }> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: `worker-clock-${Math.random()}`, profile: 'D' })
  const source = createPagedWorkbookSource(session)
  try {
    await session.done
    const sheetId = source.sheets[0]!.sheetId
    const store = source.store(sheetId)
    let clockReads = 0
    const now = (): number => { clockReads++; return NOW }
    const rects = refs.map(r => { const { row, col } = rcOf(r); return { r0: row, r1: row, c0: col, c1: col } })
    const result = await store.readConditionalFormatting!({ rects, stamp: { owner: 'clock', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 8 * 1024 * 1024 },
      createCfMemory(), { signal: new AbortController().signal, now } as never)
    const states = result?.kind === 'ready' ? result.states : []
    const fills = refs.map(r => { const { row, col } = rcOf(r); return states.find(s => s.row === row && s.col === col)?.properties['fill.bg'] as string | undefined })
    return { clockReads, fills, kind: result?.kind ?? 'none' }
  } finally { await source.release() }
}

describe('the worker samples the clock only when a rule reads it (final review F3)', () => {
  it('a literal cellIs sheet paints and takes no clock sample', async () => {
    const out = await workerRead(workbookBytes(CELLS, cfBlock('B2:B4', cellIsRule('greaterThan', '4', 1, 1))), ['B2', 'B3', 'B4'])
    expect(out.kind).toBe('ready')
    expect(out.fills).toEqual([undefined, GREEN, GREEN])
    expect(out.clockReads).toBe(0)
  })

  it('a sheet with a formula rule takes exactly one clock sample per preparation (positive control)', async () => {
    await withPaint(async () => {
      const out = await workerRead(workbookBytes(CELLS, cfBlock('B2:B4', cellIsRule('greaterThan', '4', 1, 1)) + cfBlock('C2:C4', exprRule('$A2&gt;1', 2, 1))), ['B2', 'B3', 'C3'])
      expect(out.kind).toBe('ready')
      expect(out.clockReads).toBe(1)
    })
  })

  it('a timePeriod sheet (a literal rule that reads the clock) takes exactly one sample', async () => {
    const out = await workerRead(workbookBytes(CELLS, cfBlock('A2:A4', timePeriodRule(1, 1))), ['A2', 'A3'])
    expect(out.kind).toBe('ready')
    expect(out.clockReads).toBe(1)
  })
})
