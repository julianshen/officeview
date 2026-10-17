// @vitest-environment node
import { describe, expect, it, vi } from 'vitest'
import { remoteViewReplies } from './helpers/remote-view-replies'
import { viewReplyBytes } from '../src/worker/viewerRemoteSource'
import type { RowPermutationBand, SortSpec } from '../src/xlsx/view/contracts'

const N = 40000
const band = (): RowPermutationBand => ({ r0: 0, r1: N - 1, sourceAtOffset: Uint32Array.from({ length: N }, (_, i) => N - i - 1), displayOffsetOfSource: Uint32Array.from({ length: N }, (_, i) => N - i - 1) })
const spec: SortSpec = { range: { r0: 0, r1: N - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }

describe('worker-backed sort adoption ownership', () => {
  it('moves the reply buffers to the view without charging a second copy', async () => {
    const h = remoteViewReplies(), ledger = h.remote.mainLedger
    const b = band()
    const { result, out } = await h.sort(1, [b], spec)
    expect(result.kind).toBe('committed')
    expect(h.remote.view('s').order.bands[0].sourceAtOffset).toBe(b.sourceAtOffset)
    expect(h.remote.view('s').order.bands[0].displayOffsetOfSource).toBe(b.displayOffsetOfSource)
    expect(ledger.leaseOf(out)).toBeUndefined()
    expect(ledger.peakBytes).toBeLessThanOrEqual(viewReplyBytes(N) + 4096)
    const view = h.remote.view('s'), visibility = view.visibility!
    expect(ledger.snapshot().byPool.viewBuffers).toBe(N * 8 + visibility.words.byteLength + visibility.summaries.byteLength)
  })

  it('fallback rebuilds an invalid inverse and ends staging before the new view charge', async () => {
    const h = remoteViewReplies(), ledger = h.remote.mainLedger
    await h.sort(1, [band()])
    const previous = ledger.snapshot().byPool.viewBuffers
    let peak = previous
    const set = ledger.set.bind(ledger)
    vi.spyOn(ledger, 'set').mockImplementation((...args) => { set(...args); peak = Math.max(peak, ledger.snapshot().byPool.viewBuffers) })
    const b = band()
    b.displayOffsetOfSource.fill(0) // valid forward permutation, unusable wire inverse
    const { result } = await h.sort(2, [b])
    expect(result.kind).toBe('committed')
    const view = h.remote.view('s')
    expect(view.order.kind).toBe('bands')
    expect(view.order.bands[0].sourceAtOffset.buffer).not.toBe(b.sourceAtOffset.buffer)
    expect(view.order.sourceToDisplay(123)).toBe(N - 124)
    expect(ledger.ownerBytes('viewBuffers', 'staging:s')).toBe(0)
    const held = N * 8 + view.visibility!.words.byteLength + view.visibility!.summaries.byteLength
    expect(ledger.snapshot().byPool.viewBuffers).toBe(held)
    // No display keeps the previous view here: held contains only the new view,
    // so this bound has no spare previous-view bytes to hide a staging overlap.
    expect(peak).toBeLessThanOrEqual(previous + held)
  })

  it('settles an older reply immediately after a newer sort has committed', async () => {
    const h = remoteViewReplies(), owner = h.remote.identity('s'), signal = new AbortController().signal
    const older = h.remote.sort(owner, null, signal), oldCommand = await h.next('sort')
    const newer = h.remote.sort(owner, null, signal), newCommand = await h.next('sort')
    h.reply(newCommand, { kind: 'view', delta: h.delta(2, [band()]) })
    expect((await newer).kind).toBe('committed')
    const before = h.remote.mainLedger.snapshot()
    const out = h.reply(oldCommand, { kind: 'view', delta: h.delta(1, [band()]) })
    expect(await older).toEqual({ kind: 'stale' })
    expect(h.remote.identity('s').viewGen).toBe(2)
    expect(h.remote.mainLedger.leaseOf(out)).toBeUndefined()
    expect(h.remote.mainLedger.snapshot().byPool.replies).toBe(before.byPool.replies)
  })
})


it('copies padded wire views so no uncharged backing-store tails remain in the committed view', async () => {
  const h = remoteViewReplies()
  const b: RowPermutationBand = { r0: 0, r1: 1, sourceAtOffset: new Uint32Array(new ArrayBuffer(1_000_000), 128, 2), displayOffsetOfSource: new Uint32Array(new ArrayBuffer(1_000_000), 256, 2) }
  b.sourceAtOffset.set([1, 0]); b.displayOffsetOfSource.set([1, 0])
  const { result, out } = await h.sort(1, [structuredClone(b)])
  expect(result.kind).toBe('committed')
  const view = h.remote.view('s'), adopted = view.order.bands[0]
  expect(adopted.sourceAtOffset.buffer.byteLength).toBe(8)
  expect(adopted.displayOffsetOfSource.buffer.byteLength).toBe(8)
  expect(view.order.displayToSource(0)).toBe(1)
  expect(view.order.sourceToDisplay(0)).toBe(1)
  expect(h.remote.mainLedger.leaseOf(out)).toBeUndefined()
  expect(h.remote.mainLedger.snapshot().byPool.viewBuffers).toBe(16 + view.visibility!.words.byteLength + view.visibility!.summaries.byteLength)
})
