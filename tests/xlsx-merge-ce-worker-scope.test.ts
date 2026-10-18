// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { paneLogicalBoxes } from '../src/xlsx/view/frozenPanes'

it('region union preserves corner/top requests, excludes the hull gap, and layoutGen cancels every old owner with credits returned', async () => {
  const owner = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
  const signals: AbortSignal[] = [], closed: number[] = []
  const remote = {
    pageLimits: { tile: 8, frame: 32 }, pagesInFlight: 0, pageFrame: undefined,
    beginPageFrame: RemoteWorkbook.prototype.beginPageFrame,
    admitPages: RemoteWorkbook.prototype.admitPages,
    releasePages: RemoteWorkbook.prototype.releasePages,
    merges: () => ({ gen: 0 }), identity: () => owner,
    store: () => ({ captureViewportRequest: () => { const id = closed.length; closed.push(0); return { live: () => true, close: () => { closed[id]++ } } } }),
    viewport: (_owner: unknown, _window: unknown, signal: AbortSignal) => { signals.push(signal); return new Promise(resolve => signal.addEventListener('abort', () => resolve({ kind: 'cancelled' }), { once: true })) },
    mainLedger: new MainMemoryLedger(), notify: vi.fn(), profile: 'D', isClosed: false,
  }
  Object.defineProperty(remote, 'currentPageFrame', { get: () => remote.pageFrame })
  const plans = new WorkerViewportPlans(remote as never)
  const windowOf = (x: number, y: number, width = 640, height = 400) => ({ box: { x, y, width, height }, zoom: 1, deviceScale: 1 })
  const frame = (x: number, y: number, generation: number) => plans.beginFrame(owner, windowOf(x, y), paneLogicalBoxes({ rows: 1, cols: 1, left: 64, top: 20 }, { x, y }, 1, { width: 640, height: 400 }), generation)
  const drain = () => new Promise(resolve => setTimeout(resolve, 0))
  try {
    frame(5120, 5120, 0)
    for (const [x, y] of [[0, 0], [5120, 0], [0, 5120], [2048, 2048]]) expect(plans.get(owner, windowOf(x, y, 512, 512), { width: 1e6, height: 1e6 }, true).kind).toBe('pending')
    expect(remote.pagesInFlight).toBe(32)
    frame(5120, 51200, 0)
    expect(signals.map(signal => signal.aborted)).toEqual([false, false, true, true])
    await drain()
    expect(remote.pagesInFlight).toBe(16)
    expect(closed).toEqual([0, 0, 1, 1])
    // Same spatial union, new resized-layout owner: pinned requests must not survive.
    frame(5120, 51200, 1)
    expect(signals.every(signal => signal.aborted)).toBe(true)
    await drain()
    expect(remote.pagesInFlight).toBe(0)
    expect(closed).toEqual([1, 1, 1, 1])
    plans.get(owner, windowOf(0, 0, 512, 512), { width: 1e6, height: 1e6 }, true)
    expect(signals).toHaveLength(5)
    expect(signals[4].aborted).toBe(false)
    expect(remote.pagesInFlight).toBe(8)
  } finally { plans.dispose(); await drain() }
  expect(remote.pagesInFlight).toBe(0)
  expect(closed).toEqual([1, 1, 1, 1, 1])
})
