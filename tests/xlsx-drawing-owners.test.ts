// @vitest-environment node
import { expect, it, vi } from 'vitest'
import v8 from 'node:v8'
import vm from 'node:vm'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { admitDrawingOwners, drawingOwnerOf, drawingOwnerMetadataBytes, DRAWING_OWNER_WORKSPACE_BYTES } from '../src/xlsx/view/drawingOwners'
import { remapDrawings, SIZED_DRAWING_BYTES } from '../src/xlsx/view/axisSizes'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { SheetImageCache } from '../src/components/xlsx/sheetAssets'
import type { SheetDrawings } from '../src/xlsx/view/contracts'
import type { XlsxDrawing } from '../src/xlsx/types'

v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void
async function collect() { for (let i = 0; i < 12; i++) { gc(); await new Promise(resolve => setTimeout(resolve, 10)) } }
const E = 9525, budget = 4 * 1024 * 1024
function payload(): SheetDrawings {
  return { completeness: 'complete', diagnostics: [], bytes: 0, images: [], drawings: [{
    xEmu: 80 * E, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
    source: { partPath: 'd', treePath: 'twoCellAnchor[0]/sp[0]', element: 'sp', representation: 'native' },
    group: { off: { x: 80 * E, y: 0 }, ext: { width: 80 * E, height: 20 * E }, chOff: { x: 0, y: 0 }, chExt: { width: 80 * E, height: 20 * E } },
    children: [], content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values: [{ state: 'value', v: 42 }, { state: 'value', v: 84 }] }] }] } } as never,
    image: { data: new Uint8Array(1024), mime: 'image/png' },
  }] }
}
// Two roots + node/source/group/children/content/chart + image/data/buffer =11.
const metadataBytes = 128 + 11 * 64

it('charges exactly the explicit source metadata beside one data allocation, with workspace overlap reserved', () => {
  const p = payload(), ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('drawings', [p.drawings], 8192)
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  expect(drawingOwnerMetadataBytes(p.drawings)).toBe(metadataBytes)
  expect(ledger.snapshot().byPool.drawings).toBe(8192 + metadataBytes)
  expect(ledger.peakBytes).toBe(8192 + DRAWING_OWNER_WORKSPACE_BYTES + metadataBytes)
  expect(drawingOwnerOf((p.drawings[0].content as any).chart)).toBe(p.drawings)
  expect(drawingOwnerOf(p.drawings[0].image!.data!.buffer)).toBe(p.drawings)
})

it('does not traverse or construct links before workspace admission', () => {
  const p = payload(), ledger = new MainMemoryLedger(), inspect = vi.fn(() => [])
  Object.defineProperty(p.drawings[0], 'children', { get: inspect })
  expect(admitDrawingOwners(p, ledger, DRAWING_OWNER_WORKSPACE_BYTES - 1)).toBe(false)
  expect(inspect).not.toHaveBeenCalled()
  expect(drawingOwnerOf(p.drawings[0])).toBeUndefined()
  expect(ledger.snapshot()).toMatchObject({ bytes: 0, peakBytes: 0 })
})

it('refuses retained metadata before the first link when workspace plus metadata cannot coexist', () => {
  const p = payload(), ledger = new MainMemoryLedger()
  expect(admitDrawingOwners(p, ledger, DRAWING_OWNER_WORKSPACE_BYTES + metadataBytes - 1)).toBe(false)
  expect(drawingOwnerOf(p.drawings)).toBeUndefined()
  expect(drawingOwnerOf((p.drawings[0].content as any).chart)).toBeUndefined()
  expect(drawingOwnerMetadataBytes(p.drawings)).toBe(0)
  expect(ledger.snapshot()).toMatchObject({ bytes: 0, peakBytes: DRAWING_OWNER_WORKSPACE_BYTES })
})

it('empty drawing replies under full pressure need no workspace, metadata or ownership context', () => {
  const ledger = new MainMemoryLedger(), p: SheetDrawings = { completeness: 'complete', diagnostics: [], bytes: 0, images: [], drawings: [] }
  ledger.set('dom', 'pressure', budget)
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  expect(ledger.snapshot()).toMatchObject({ bytes: budget, peakBytes: budget })
  expect(drawingOwnerOf(p.drawings)).toBeUndefined()
})

it('warm extraction reuses its admitted context without another workspace or data/metadata debit', () => {
  const ledger = new MainMemoryLedger(), p = payload()
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  const snapshot = ledger.snapshot()
  expect(admitDrawingOwners(p, ledger, snapshot.bytes)).toBe(true)
  expect(ledger.snapshot()).toEqual(snapshot)
})

it('unique drawing nodes and diagram shapes share the existing parser bound; chart points do not consume it', () => {
  const p = payload(), ledger = new MainMemoryLedger()
  const node = p.drawings[0]
  Object.defineProperty((node.content as any).chart, 'groups', { get: () => { throw Error('must not traverse chart points') } })
  p.drawings = Array(10_000).fill(node)
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  expect(drawingOwnerMetadataBytes(p.drawings)).toBe(metadataBytes)
  const diagram: SheetDrawings = { completeness: 'complete', diagnostics: [], bytes: 0, images: [], drawings: [{ content: { kind: 'diagram', shapes: Array.from({ length: 9999 }, () => ({})) } } as unknown as XlsxDrawing] }
  expect(admitDrawingOwners(diagram, ledger, budget)).toBe(true)
  const tooMany: SheetDrawings = { ...diagram, drawings: [{ content: { kind: 'diagram', shapes: Array.from({ length: 10_000 }, () => ({})) } } as unknown as XlsxDrawing] }
  expect(admitDrawingOwners(tooMany, ledger, budget)).toBe(false)
  expect(drawingOwnerOf(tooMany.drawings)).toBeUndefined()
})

it('repeated remaps flatten source ownership while a detached copied group keeps only its own copy allocation', async () => {
  let p: SheetDrawings | null = payload()
  const ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('drawings', [p.drawings], 8192)
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  const baseBytes = 8192 + metadataBytes
  const rows = createAxisMetrics({ logicalCount: 4, defaultSizeCss: 20 })
  const axis = (first: number) => ({ rows, cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 80, overrides: [{ start: 0, end: 0, sizeCss: first }] }) })
  const remap = (list: readonly XlsxDrawing[], from: number, to: number) => {
    ledger.set('drawings', 'copy-work', SIZED_DRAWING_BYTES)
    const copies = remapDrawings(list, axis(from), axis(to))
    ledger.adoptWhileReachable('drawings', 'copy-work', 'drawings', [copies], SIZED_DRAWING_BYTES)
    return copies
  }
  let first: XlsxDrawing[] | null = remap(p.drawings, 80, 120)
  const weakFirst = new WeakRef(first)
  let second: XlsxDrawing[] | null = remap(first, 120, 160)
  expect(second[0].xEmu / E).toBe(160)
  expect(drawingOwnerOf(second[0])).toBe(p.drawings)
  first = null; await collect()
  expect(weakFirst.deref()).toBeUndefined()
  expect(ledger.snapshot().byPool.drawings).toBe(baseBytes + SIZED_DRAWING_BYTES)
  let group: object | null = second[0].group!
  second = null; await collect()
  expect(drawingOwnerOf(group)).toBe(p.drawings)
  expect(ledger.snapshot().byPool.drawings).toBe(baseBytes + SIZED_DRAWING_BYTES)
  group = null; await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(baseBytes)
  p = null; await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(0)
})

it('the actual image cache key retains raw image data until its final release without a duplicate data charge', async () => {
  const ledger = new MainMemoryLedger(), cache = new SheetImageCache(64)
  let p: SheetDrawings | null = payload()
  ledger.chargeWhileReachable('drawings', [p.drawings], 8192)
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  const asset = new WeakRef(p.drawings[0].image!)
  await cache.acquire(p.drawings[0].image!, p.drawings[0].image!, 4, async () => ({ width: 1, height: 1 }) as CanvasImageSource)
  p = null; await collect()
  expect(asset.deref()!.data).toHaveLength(1024)
  expect(ledger.snapshot().byPool.drawings).toBe(8192 + metadataBytes)
  cache.release(asset.deref()!)
  await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(0)
  expect(cache.chargedBytes).toBe(0)
})

// Indexed media are aliases, not parser drawing nodes. This fixture intentionally
// shares the one media object; allowance duplicates over-count conservatively.
it('does not invent a 10000-indexed-image cap: preprices aliases against the unchanged byte budget', () => {
  const p = payload(), ledger = new MainMemoryLedger()
  p.images = Array(10001).fill({ data: p.drawings[0].image!.data!, mime: 'image/png' })
  expect(admitDrawingOwners(p, ledger, 8 * 1024 * 1024)).toBe(true)
  expect(drawingOwnerMetadataBytes(p.drawings)).toBe(metadataBytes + 10001 * 6 * 64)
})

it('counts unique nodes rather than rejecting a long repeated-reference list', () => {
  const p = payload(), ledger = new MainMemoryLedger()
  p.drawings = Array(10001).fill(p.drawings[0])
  expect(admitDrawingOwners(p, ledger, budget)).toBe(true)
  expect(drawingOwnerMetadataBytes(p.drawings)).toBe(metadataBytes)
})
