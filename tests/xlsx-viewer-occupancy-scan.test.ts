/**
 * S10.2: paged range detection only needs which cells are populated. The
 * paged store's occupancy scan yields exactly scanPopulated's coordinates, in
 * the same order and page boundaries, without decoding payloads (in Chrome the
 * decode made a 1M-row header sort's admission take ~1.5 s inside the worker).
 */
import { describe, expect, it } from 'vitest'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { colName, xlsxBytes } from './helpers/xlsx-package'

const quiet = { yield: async () => {} }
const PAGED = { object: { rows: 1 } }
async function open(bytes: Uint8Array) {
  const kernel = new ViewerKernel({ parse: async b => parseXlsx(await OfficePackage.load(b)), importLimits: PAGED }, quiet)
  const replies: WorkerReply[] = []
  await kernel.handle({ kind: 'open', requestId: 'o', docId: 'd', source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' } as WorkerCommand, r => { replies.push(r) })
  return kernel.session('d')!.source.store('1')
}
const xml = () => {
  const rows: string[] = []
  for (let r = 0; r < 300; r++) {
    const cells: string[] = []
    for (let c = 0; c < 6; c++) {
      if ((r * 5 + c) % 7 === 0) continue // gaps
      const ref = `${colName(c)}${r + 1}`
      if (c === 5 && r % 9 === 0) cells.push(`<c r="${ref}"><f>A${r + 1}*2</f></c>`) // formula without cached value: unavailable, still populated
      else if (c === 4 && r % 11 === 0) cells.push(`<c r="${ref}" s="0"/>`) // style-only: not populated
      else cells.push(c % 2 ? `<c r="${ref}" t="inlineStr"><is><t>t${r}-${c}</t></is></c>` : `<c r="${ref}"><v>${r * c}</v></c>`)
    }
    rows.push(`<row r="${r + 1}">${cells.join('')}</row>`)
  }
  return rows.join('')
}

describe('paged occupancy scan', () => {
  it('matches scanPopulated coordinates and page boundaries for every rect and limit', async () => {
    const store = await open(xlsxBytes({ sheetData: xml() }))
    expect(typeof store.scanOccupied).toBe('function')
    const opts = () => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
    for (const range of [{ r0: 0, r1: 299, c0: 0, c1: 5 }, { r0: 17, r1: 233, c0: 2, c1: 4 }, { r0: 299, r1: 1048575, c0: 0, c1: 16383 }, { r0: 5, r1: 5, c0: 5, c1: 5 }]) {
      for (const limit of [1, 7, 4096]) {
        const full: { row: number; col: number }[][] = [], occ: { row: number; col: number }[][] = []
        const flags: string[] = [], occFlags: string[] = []
        for await (const page of store.scanPopulated({ range, limit }, opts())) { full.push(page.items.map(c => ({ row: c.row, col: c.col }))); flags.push(page.completeness) }
        for await (const page of store.scanOccupied!({ range, limit }, opts())) { occ.push(page.items.map(c => ({ row: c.row, col: c.col }))); occFlags.push(page.completeness) }
        expect(occ, `${JSON.stringify(range)} limit ${limit}`).toEqual(full)
        expect(occFlags).toEqual(flags)
      }
    }
  })
})
