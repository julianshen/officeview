// @vitest-environment node
/** The pacing yield (spec §10.2): a real macrotask, never a clamped 0 ms timer
 * in a browser-like host, one shared channel, FIFO with queued tasks. */
import { describe, expect, it, vi } from 'vitest'
import { macrotaskYield } from '../src/xlsx/view/scheduler'

/** A browser-like host: no setImmediate, no Node process; a timer spy. */
function browserHost() {
  const setTimeout = vi.fn((callback: () => void) => { globalThis.setTimeout(callback, 0) })
  return { host: { MessageChannel, setTimeout }, setTimeout }
}
async function microtaskChain(length: number): Promise<number> {
  let done = 0
  for (let i = 0; i < length; i++) { await Promise.resolve(); done++ }
  return done
}

describe('macrotaskYield', () => {
  it('in a browser-like host yields through one shared MessageChannel, not a timer', async () => {
    const { host, setTimeout } = browserHost()
    const construct = vi.fn()
    class CountedChannel extends MessageChannel { constructor() { super(); construct() } }
    const yieldOnce = macrotaskYield({ ...host, MessageChannel: CountedChannel })
    for (let i = 0; i < 50; i++) await yieldOnce()
    expect(setTimeout).not.toHaveBeenCalled()
    expect(construct).toHaveBeenCalledTimes(1)
  })
  it('is a macrotask: microtasks queued before it all run first', async () => {
    const { host } = browserHost()
    const yieldOnce = macrotaskYield(host)
    let chain = 0
    const microtasks = microtaskChain(1000).then(done => { chain = done })
    await yieldOnce()
    expect(chain).toBe(1000)
    await microtasks
  })
  it('services a cancel message queued before it (browser posted-message FIFO)', async () => {
    // A browser delivers posted messages from one task source in order, one
    // task each; model that so the check does not depend on Node's per-port queues.
    const fifo: (() => void)[] = []
    const pump = () => { const next = fifo.shift(); next?.(); if (fifo.length) setTimeout(pump, 0) }
    class FifoChannel {
      port1: { onmessage: ((event: unknown) => void) | null; close(): void } = { onmessage: null, close() {} }
      port2 = { postMessage: (data: unknown) => { const port = this.port1; fifo.push(() => port.onmessage?.({ data })); if (fifo.length === 1) setTimeout(pump, 0) }, close() {} }
    }
    const yieldOnce = macrotaskYield({ MessageChannel: FifoChannel as unknown as typeof MessageChannel, setTimeout: vi.fn() })
    const controller = new AbortController()
    const cancel = new FifoChannel()
    cancel.port1.onmessage = () => controller.abort()
    cancel.port2.postMessage(null)
    await yieldOnce()
    expect(controller.signal.aborted).toBe(true)
  })
  it('resumes concurrent waiters in order, one per yield', async () => {
    const { host } = browserHost()
    const yieldOnce = macrotaskYield(host)
    const order: number[] = []
    await Promise.all([0, 1, 2, 3].map(async i => { await yieldOnce(); order.push(i) }))
    expect(order).toEqual([0, 1, 2, 3])
  })
  it('uses the native setImmediate on Node, and a timer only where nothing else exists', async () => {
    const immediate = vi.fn((callback: () => void) => { setImmediate(callback) })
    const timer = vi.fn((callback: () => void) => { setTimeout(callback, 0) })
    await macrotaskYield({ process: { versions: { node: '22' } }, setImmediate: immediate, MessageChannel, setTimeout: timer })()
    expect(immediate).toHaveBeenCalledTimes(1)
    await macrotaskYield({ setTimeout: timer })()
    expect(timer).toHaveBeenCalledTimes(1)
    // A setImmediate without Node (a bundled polyfill) is not trusted.
    const polyfill = vi.fn((callback: () => void) => { setImmediate(callback) })
    await macrotaskYield({ setImmediate: polyfill, MessageChannel, setTimeout: timer })()
    expect(polyfill).not.toHaveBeenCalled()
  })
})
