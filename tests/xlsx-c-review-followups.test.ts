/** Stream C review follow-ups that the reviewer's reproductions do not cover directly. */
import { pts } from './helpers/chart-points'
import { describe, expect, it } from 'vitest'
import { RemoteDrawingCache, type DrawingHost } from '../src/worker/viewerRemoteDrawings'
import type { CacheIdentity, SheetDrawings, WorkerReply } from '../src/xlsx/view/contracts'
import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { lineToViewport, paneLogicalBoxes } from '../src/xlsx/view/frozenPanes'

const identity = (sheetId: string): CacheIdentity => ({ docId: 'D', sheetId, valueGen: 0, fontGen: 0, viewGen: 0 })
const tick = () => new Promise(resolve => setTimeout(resolve, 0))

describe('charts waiting on another sheet', () => {
  it('a payload with pending sheets is read again once those sheets complete', async () => {
    const pending: Array<(reply: WorkerReply<'read'>) => void> = []
    const states = new Map<string, 'complete' | 'loading'>([['1', 'complete'], ['2', 'loading']])
    const h: DrawingHost = {
      identity, loadState: id => ({ completeness: states.get(id) ?? 'complete' }), nextRequestId: () => `r${pending.length}`,
      request: () => new Promise(resolve => pending.push(resolve as never)), changed: () => {},
    }
    const reply = (result: SheetDrawings) => pending.at(-1)!({ commandKind: 'read', requestId: 'x', owner: identity('1'), diagnostics: [], kind: 'complete', result: { kind: 'read', result: { op: 'drawings', result } } } as WorkerReply<'read'>)
    const cache = new RemoteDrawingCache(h, 1e6)
    cache.get('1')
    reply({ completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0, pendingSheets: ['2'] })
    await tick()
    expect(cache.get('1').state).toBe('ready')
    expect(pending).toHaveLength(1)
    states.set('2', 'complete')
    cache.structureComplete('2')
    expect(pending).toHaveLength(2)
    reply({ completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0 })
    await tick()
    expect(cache.get('1').state).toBe('ready')
    expect(pending).toHaveLength(2)
  })
})

describe('resident charts follow the document value policy', () => {
  const xml = '<chartSpace><chart><plotArea><barChart><ser><idx val="0"/><val><numRef><f>Styled!B1:B2</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
  const patch = async (zip: import('jszip')) => {
    const path = 'xl/worksheets/sheet1.xml', sheet = await zip.file(path)!.async('string')
    zip.file(path, sheet.replace('<v>99</v>', '<f ca="1">A1+4</f><v>99</v>').replace('<v>50</v>', '<f>A1*2</f>'))
  }
  const cells = [{ ref: 'A1', v: 3 }, { ref: 'B1', v: 99 }, { ref: 'B2', v: 50 }]
  it('engine: the evaluated values', async () => {
    const doc = await parseXlsx(await chartWorkbook(xml, { cells, patch }))
    const chart = doc.sheets[0].drawings![0].content
    expect(chart?.kind === 'xlsxChart' && chart.chart.groups[0].series[0].values).toEqual(pts(7, 6))
  })
  it('cached: the saved value; an uncached formula has no value', async () => {
    const doc = await parseXlsx(await chartWorkbook(xml, { cells, patch }), { valuePolicy: 'cached' })
    const chart = doc.sheets[0].drawings![0].content
    expect(chart?.kind === 'xlsxChart' && chart.chart.groups[0].series[0].values).toEqual(pts(99, 'unavailable'))
  })
})

describe('pane-aware grid lines (for resize handles)', () => {
  const frozen = { rows: 1, cols: 2, left: 128, top: 20 }, viewport = { width: 640, height: 400 }
  it('maps a header line per region and hides lines under the strip or past the pane', () => {
    const origin = { x: 1000, y: 500 }
    expect(lineToViewport(frozen, origin, 1, viewport, 'x', 64)).toBe(64)
    expect(lineToViewport(frozen, origin, 1, viewport, 'x', 128)).toBe(128)
    expect(lineToViewport(frozen, origin, 1, viewport, 'x', 1064)).toBeNull()
    expect(lineToViewport(frozen, origin, 1, viewport, 'x', 1192)).toBe(192)
    expect(lineToViewport(frozen, origin, 2, viewport, 'y', 520)).toBe(40)
    expect(lineToViewport(frozen, origin, 1, viewport, 'y', 510)).toBeNull()
    expect(lineToViewport(frozen, origin, 1, viewport, 'x', 2000)).toBeNull()
    expect(lineToViewport(undefined, origin, 1, viewport, 'x', 1064)).toBe(64)
  })
  it('logical boxes of the regions are disjoint, never their hull', () => {
    expect(paneLogicalBoxes(frozen, { x: 1000, y: 500 }, 1, viewport)).toEqual([
      { x: 1128, y: 520, width: 512, height: 380 }, { x: 1128, y: 0, width: 512, height: 20 },
      { x: 0, y: 520, width: 128, height: 380 }, { x: 0, y: 0, width: 128, height: 20 },
    ])
  })
})
