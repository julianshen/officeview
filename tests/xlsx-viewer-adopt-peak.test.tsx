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

  it('settles a stale keyed reply whose reservation already moved into its lease', async () => {
    const h = remoteViewReplies(), ledger = h.remote.mainLedger, owner = h.remote.identity('s'), signal = new AbortController().signal
    const baseline = ledger.snapshot().byPool.replies
    const older = h.remote.sort(owner, spec, signal)
    h.reply(await h.next('prepareColumns'), { kind: 'prepareColumns', cols: new Uint16Array([0]) })
    const oldCommand = await h.next('sort')
    const newer = h.remote.sort(owner, spec, signal)
    h.reply(await h.next('prepareColumns'), { kind: 'prepareColumns', cols: new Uint16Array([0]) })
    const newCommand = await h.next('sort')
    // Both reservations are live; the prepare replies keep their own small leases.
    const unreserved = ledger.snapshot().byPool.replies - 2 * viewReplyBytes(N)
    expect(unreserved).toBeGreaterThanOrEqual(baseline)
    h.reply(newCommand, { kind: 'view', delta: h.delta(2, [band()]) })
    expect((await newer).kind).toBe('committed')
    const out = h.reply(oldCommand, { kind: 'view', delta: h.delta(1, [band()]) })
    expect(await older).toEqual({ kind: 'stale' })
    expect(h.remote.identity('s').viewGen).toBe(2)
    expect(ledger.leaseOf(out)).toBeUndefined()
    expect(ledger.snapshot().byPool.replies).toBe(unreserved)
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

  it('keeps every local typed array of a padded-input fallback charged while it is staged', async () => {
    const h = remoteViewReplies(), ledger = h.remote.mainLedger, n = N
    const padded = (offset: number) => new Uint32Array(new ArrayBuffer(n * 4 + 1024), offset, n)
    const b: RowPermutationBand = { r0: 0, r1: n - 1, sourceAtOffset: padded(128), displayOffsetOfSource: padded(256) }
    for (let i = 0; i < n; i++) { b.sourceAtOffset[i] = n - i - 1; b.displayOffsetOfSource[i] = n - i - 1 }
    // From the reply until staging ends, every fresh Uint32Array (constructed or
    // sliced; a view over an existing buffer owns no bytes) must be covered by
    // the staging charge plus what is left of the sort's reservation.
    let reservation = '', live = 0, uncovered = 0, counting = false, staged = false
    const charged = () => ledger.ownerBytes('viewBuffers', 'staging:s') + (reservation ? ledger.ownerBytes('replies', reservation) : 0)
    const count = (array: Uint32Array) => {
      if (!counting) return
      live += array.byteLength
      uncovered = Math.max(uncovered, live - charged())
    }
    const Native = globalThis.Uint32Array, slice = Native.prototype.slice
    globalThis.Uint32Array = new Proxy(Native, { construct(target, args) {
      const array = Reflect.construct(target, args) as Uint32Array
      if (!(args[0] instanceof ArrayBuffer)) count(array)
      return array
    } })
    Native.prototype.slice = function (this: Uint32Array, ...args: Parameters<Uint32Array['slice']>) {
      const array = slice.apply(this, args)
      count(array)
      return array
    }
    // release() is set(…, 0): the staging window is the positive set until it.
    const set = ledger.set.bind(ledger)
    vi.spyOn(ledger, 'set').mockImplementation((pool, owner, bytes) => {
      if (owner.startsWith('view-reply:') && bytes > 0) reservation = owner
      if (owner === 'staging:s') { if (bytes > 0) staged = true; else counting = false }
      set(pool, owner, bytes)
    })
    const sortSpec: SortSpec = { ...spec, range: { ...spec.range, r1: n - 1 } }, wire = structuredClone(b)
    try {
      counting = true
      const { result } = await h.sort(1, [wire], sortSpec)
      expect(result.kind).toBe('committed')
    } finally {
      globalThis.Uint32Array = Native
      Native.prototype.slice = slice
    }
    const committed = h.remote.view('s').order.bands[0]
    expect(committed.sourceAtOffset.byteLength + committed.displayOffsetOfSource.byteLength).toBe(n * 8)
    expect(reservation).not.toBe('')
    expect(staged).toBe(true)
    expect(live).toBeGreaterThanOrEqual(n * 8)
    expect(uncovered).toBe(0)
  })
})
