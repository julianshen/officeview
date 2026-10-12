/**
 * S13.2: the main-thread drawing cache of worker-backed sheets adopts only the
 * reply it waits for from the sheet's current owner, retries stale/cancelled
 * reads a bounded number of times, turns failures into visible diagnostics and
 * keeps retained payloads within its byte budget. The source-side payload
 * charges pictures against the transfer budget and never mutates the resident
 * model it was cloned from.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { RemoteDrawingCache, documentWithDrawings, drawingNotices, sheetWithDrawings, type DrawingHost } from '../src/worker/viewerRemoteDrawings'
import { sheetDrawingPayload } from '../src/xlsx/view/drawingPayload'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession } from '../src/worker/viewerSession'
import type { CacheIdentity, SheetDrawings, WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument, XlsxDrawing, XlsxSheet } from '../src/xlsx/types'

afterEach(() => { vi.useRealTimers() })

const node = (name: string, x: number, data?: Uint8Array): XlsxDrawing => ({
  source: { partPath: 'xl/drawings/drawing1.xml', treePath: name, element: data ? 'pic' : 'sp', name, representation: 'native' },
  xEmu: x * 9525, yEmu: 0, widthEmu: 10 * 9525, heightEmu: 10 * 9525,
  ...(data ? { image: { data, mime: 'image/png', partPath: `xl/media/${name}.png` } } : {}),
})
const payload = (drawings: XlsxDrawing[], bytes = 0): SheetDrawings => ({ completeness: 'complete', drawings, images: [], diagnostics: [], bytes })
const identity = (sheetId: string, valueGen = 0): CacheIdentity => ({ docId: 'D', sheetId, valueGen, fontGen: 0, viewGen: 0 })

/** A host whose replies the test resolves by hand. */
function host(complete = true) {
  const pending: Array<{ owner: CacheIdentity; resolve: (reply: WorkerReply<'read'> | symbol) => void; signal?: AbortSignal }> = []
  const changed = vi.fn()
  const owners = new Map<string, CacheIdentity>()
  const h: DrawingHost = {
    identity: sheetId => owners.get(sheetId) ?? identity(sheetId),
    loadState: () => ({ completeness: complete ? 'complete' : 'loading' }),
    nextRequestId: () => `r${pending.length}`,
    request: (command, signal) => new Promise(resolve => pending.push({ owner: command.owner, resolve, signal })),
    changed,
  }
  const reply = (index: number, result: SheetDrawings, owner = pending[index].owner) => pending[index].resolve({ commandKind: 'read', requestId: 'x', owner, diagnostics: [], kind: 'complete', result: { kind: 'read', result: { op: 'drawings', result } } } as WorkerReply<'read'>)
  return { h, pending, changed, owners, reply }
}
const tick = () => new Promise(resolve => setTimeout(resolve, 0))

describe('RemoteDrawingCache', () => {
  it('waits for the structure, then adopts the drawings and bumps the generation', async () => {
    let complete = false
    const t = host()
    t.h.loadState = () => ({ completeness: complete ? 'complete' : 'loading' })
    const cache = new RemoteDrawingCache(t.h, 1000)
    expect(cache.get('1').state).toBe('waiting')
    expect(t.pending).toHaveLength(0)
    complete = true
    cache.structureComplete('1')
    expect(t.pending).toHaveLength(1)
    t.reply(0, payload([node('a', 0)]))
    await tick()
    const content = cache.get('1')
    expect(content.state).toBe('ready')
    expect(content.gen).toBe(1)
    expect(content.drawingsIn(0, 0, 100, 100, 100)?.map(d => d.source.name)).toEqual(['a'])
    expect(t.changed).toHaveBeenCalledWith('1')
  })

  it('never adopts a reply for another owner or a source generation that moved on; it asks again', async () => {
    vi.useFakeTimers()
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 1000)
    cache.get('1')
    t.reply(0, payload([node('foreign', 0)]), identity('1', 7))
    await vi.advanceTimersByTimeAsync(0)
    expect(cache.get('1').state).toBe('waiting')
    expect(cache.get('1').drawingsIn(0, 0, 100, 100, 100)).toEqual([])
    await vi.advanceTimersByTimeAsync(300)
    expect(t.pending).toHaveLength(2)
    // The source owner moves on while the second read is in flight.
    t.owners.set('1', identity('1', 1))
    t.reply(1, payload([node('old', 0)]), identity('1', 0))
    await vi.advanceTimersByTimeAsync(300)
    expect(cache.get('1').state).not.toBe('ready')
    expect(t.pending).toHaveLength(3)
    expect(t.pending[2].owner.valueGen).toBe(1)
    t.reply(2, payload([node('current', 0)]))
    await vi.advanceTimersByTimeAsync(0)
    expect(cache.get('1').drawingsIn(0, 0, 100, 100, 100)?.map(d => d.source.name)).toEqual(['current'])
  })

  it('retries a stale or cancelled read a bounded number of times, then reports it', async () => {
    vi.useFakeTimers()
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 1000)
    cache.get('1')
    for (let attempt = 0; attempt < 20; attempt++) {
      t.pending[attempt].resolve({ commandKind: 'read', requestId: 'x', owner: identity('1'), diagnostics: [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }], kind: 'error', code: 'stale-source', retryable: true } as WorkerReply<'read'>)
      await vi.advanceTimersByTimeAsync(300)
    }
    expect(t.pending).toHaveLength(20)
    expect(cache.get('1').state).toBe('failed')
    expect(drawingNotices(cache.get('1'))[0]).toMatch(/Drawings could not be loaded: Command owner/)
    // Transient exhaustion is not final: showing the sheet later asks once more (rate-limited).
    expect(t.pending).toHaveLength(20)
    await vi.advanceTimersByTimeAsync(5000)
    cache.get('1')
    expect(t.pending).toHaveLength(21)
    t.reply(20, payload([node('late', 0)]))
    await vi.advanceTimersByTimeAsync(0)
    expect(cache.get('1').state).toBe('ready')
    expect(drawingNotices(cache.get('1'))).toEqual([])
  })

  it('keeps the source diagnostics of a non-retryable incomplete result', async () => {
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 1000)
    cache.get('1')
    t.reply(0, { completeness: 'incomplete', drawings: [], images: [], bytes: 0, diagnostics: [{ code: 'drawings-unavailable', message: 'Drawings could not be read: boom' }], error: { code: 'parse-error', retryable: false, message: 'boom' } })
    await tick()
    expect(cache.get('1').state).toBe('failed')
    expect(drawingNotices(cache.get('1'))).toEqual(['Drawings could not be read: boom'])
    // A deterministic failure is not asked again.
    expect(t.pending).toHaveLength(1)
  })

  it('ignores replies after close or reset', async () => {
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 1000)
    cache.get('1')
    cache.close()
    t.reply(0, payload([node('late', 0)]))
    await tick()
    expect(t.changed).not.toHaveBeenCalled()
  })

  it('gives back the least recently used payloads beyond its budget', async () => {
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 100)
    cache.get('A'); t.reply(0, payload([node('a', 0)], 60)); await tick()
    cache.get('B'); t.reply(1, payload([node('b', 0)], 30)); await tick()
    expect(cache.retainedBytes).toBe(90)
    cache.get('A') // A becomes the most recent
    cache.get('C'); t.reply(2, payload([node('c', 0)], 50)); await tick()
    // B (least recent) is given back first; A + C = 110 still exceeds, so A goes too.
    expect(cache.get('C').state).toBe('ready')
    expect(cache.retainedBytes).toBeLessThanOrEqual(100)
    expect(t.changed).toHaveBeenCalledWith('B')
  })

  it('derives a stable sheet and lease document carrying the content', async () => {
    const t = host()
    const cache = new RemoteDrawingCache(t.h, 1000)
    cache.get('1')
    const picture = node('p', 20, new Uint8Array([1, 2, 3]))
    t.reply(0, sheetDrawingPayload([node('a', 0), picture], { fonts: {} } as never, [], 1000))
    await tick()
    const content = cache.get('1')
    const base: XlsxSheet = { name: 'S', sheetId: '1', rows: [], cols: [], merges: [], mergeRanges: [] }
    const sheet = sheetWithDrawings(base, content)
    expect(sheet).toBe(sheetWithDrawings(base, content))
    expect(sheet.drawings).toHaveLength(2)
    expect(sheet.drawingTheme).toBeDefined()
    const doc = documentWithDrawings({ sheets: [base] } as XlsxDocument, 0, content)
    expect(doc.sheets[0]).toBe(sheet)
    expect([...doc.images![sheet.drawings![1].imageIndex!].data]).toEqual([1, 2, 3])
  })
})

describe('source drawing payloads', () => {
  const resident = (): XlsxDocument => {
    const picture = node('p', 20, new Uint8Array(64))
    picture.imageIndex = 5 // a workbook-wide index from the resident parse
    return { sheets: [{ name: 'S', sheetId: '1', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [node('a', 0), picture] }] }
  }

  it('the object adapter clones (resident model untouched) and re-indexes per sheet', async () => {
    const doc = resident()
    const store = createObjectWorkbookSource(doc, 'D').store('1')
    const result = await store.readDrawings!({ signal: new AbortController().signal, stamp: store.stamp }, 'D')
    expect(result.completeness).toBe('complete')
    expect(result.drawings[1].imageIndex).toBe(0)
    expect(result.images).toHaveLength(1)
    expect(result.bytes).toBe(64)
    expect(result.drawings[1].image!.data).not.toBe(doc.sheets[0].drawings![1].image!.data)
    expect(doc.sheets[0].drawings![1].imageIndex).toBe(5)
  })

  it('withdraws what exceeds the budget and names it', async () => {
    const doc = resident()
    const store = createObjectWorkbookSource(doc, 'D', { drawingPayloadBytes: 10 }).store('1')
    const result = await store.readDrawings!({ signal: new AbortController().signal, stamp: store.stamp }, 'D')
    expect(result.drawings).toHaveLength(2)
    expect(result.drawings[1].image!.data).toBeUndefined()
    expect(result.drawings[1].imageIndex).toBeUndefined()
    expect(result.diagnostics.map(d => d.code)).toEqual(['drawing-asset-refused'])
    expect(doc.sheets[0].drawings![1].image!.data).toHaveLength(64)
  })

  it('a session over a store without drawing support answers with a visible diagnostic', async () => {
    const source = createObjectWorkbookSource(resident(), 'D')
    // A store of a source without drawing support (the method is optional).
    Object.defineProperty(source.store('1'), 'readDrawings', { value: undefined })
    const session = new DocumentSession('D', source)
    const out = await session.read(identity('1'), { op: 'drawings' }, new AbortController().signal)
    expect('op' in out && out.op === 'drawings' && out.result.diagnostics.map(d => d.code)).toEqual(['drawings-unavailable'])
  })
})
