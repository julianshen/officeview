// @vitest-environment node
import { expect, it } from 'vitest'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import type { ViewportCell } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

async function harness(cells: { row: number; col: number; value: string | number }[]) {
  const doc = workbookOf([{ sheetId: 's', cells }]), source = createObjectWorkbookSource(doc, 'linear-merge')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'linear-merge', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId: 'linear-merge', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], layouts: new Map() })
  return { remote, store: remote.store('s') }
}

function incoming(col: number, value: string | null, row = 0, unavailable = false): ViewportCell {
  const common = {
    source: { row, col }, bounds: { x: col * 64, y: row * 20, width: 64, height: 20 },
    cell: { id: row * 16384 + col, row, col, value, styleIndex: 0, ...(unavailable ? { formula: 'NOW()', hasCachedValue: false } : {}) },
  }
  return unavailable ? { ...common, kind: 'unavailable', reason: 'formula-value-unavailable' } : { ...common, kind: 'cell' }
}

it('merges sorted incoming cells with linear column reads as an existing row doubles', async () => {
  const reads: number[] = []
  for (const count of [128, 256]) {
    const h = await harness(Array.from({ length: count }, (_, i) => ({ row: 0, col: i * 2, value: i })))
    try {
      await h.store.load([0], 0, count * 2 - 1)
      let columnReads = 0
      // Observe the actual cached payload, including scans by staging and its
      // comparator; setup and output assertions run outside the measurement.
      for (const cell of h.store.cellsOf(0)!) {
        const col = cell.col
        Object.defineProperty(cell, 'col', { configurable: true, enumerable: true, get: () => { columnReads++; return col } })
      }
      const token = h.store.stageViewport(Array.from({ length: count }, (_, i) => incoming(i * 2 + 1, `new-${i}`)), new AbortController().signal)
      const measuredReads = columnReads
      expect(token).toBeDefined()
      try { expect(token!.commit(() => {})).toBe(true) } finally { token!.close() }
      expect(h.store.cellsOf(0)?.map(cell => cell.col)).toEqual(Array.from({ length: count * 2 }, (_, i) => i))
      reads.push(measuredReads)
      expect(measuredReads, `column reads for ${count} old + ${count} incoming cells`).toBeLessThanOrEqual(10 * (count + count))
    } finally { await h.remote.release() }
  }
  expect(reads[1], 'doubling both inputs must keep column reads linear').toBeLessThanOrEqual(reads[0] * 2.25 + 32)
})

it('merges reordered duplicate coordinates without losing old cells or widening chunk knownness', async () => {
  const h = await harness([{ row: 0, col: 0, value: 'old-zero' }, { row: 0, col: 2, value: 'old-two' }, { row: 0, col: 6, value: 'old-six' }, { row: 0, col: 600, value: 'old-tail' }, { row: 1, col: 0, value: 'other-row' }])
  try {
    await h.store.load([0, 1], 0, 255)
    await h.store.load([0], 512, 767)
    const old = h.store.cellsOf(0)!, untouched = old.find(cell => cell.col === 6), other = h.store.cellsOf(1)
    const cells = [incoming(300, 'first'), incoming(2, 'superseded'), incoming(1, 'inserted'), incoming(300, null, 0, true), incoming(0, 'zero-replaced')]
    const token = h.store.stageViewport(cells, new AbortController().signal)
    expect(token).toBeDefined()
    expect(h.store.cellsOf(0)).toBe(old)
    try { expect(token!.commit(() => {})).toBe(true) } finally { token!.close() }
    const merged = h.store.cellsOf(0)!
    expect(merged.map(cell => [cell.col, cell.value])).toEqual([[0, 'zero-replaced'], [1, 'inserted'], [2, 'superseded'], [6, 'old-six'], [300, null], [600, 'old-tail']])
    expect(merged.find(cell => cell.col === 6)).toBe(untouched)
    expect(h.store.cellsOf(1)).toBe(other)
    expect(merged.find(cell => cell.col === 300)).toMatchObject({ ref: 'KO1', unavailable: true, formula: 'NOW()', hasCachedValue: false })
    expect(h.store.isKnown({ row: 0, col: 255 })).toBe(true)
    expect(h.store.isKnown({ row: 0, col: 256 })).toBe(false)
    expect(h.store.isKnown({ row: 0, col: 300 })).toBe(false)
    expect(h.store.isKnown({ row: 0, col: 767 })).toBe(true)
    expect(h.store.cellAt({ row: 0, col: 300 })).toMatchObject({ unavailable: true, value: null })
    expect(cells[3]).toMatchObject({ kind: 'unavailable', cell: { value: null } })
  } finally { await h.remote.release() }
})
