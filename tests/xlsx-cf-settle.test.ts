// @vitest-environment node
/**
 * Task 11 fix round 1, finding 3: a worker read that never settles (no
 * failure, no notification) must not hang the preparation forever. The
 * settle waiter resolves on timeout so the caller can drop the stuck entry
 * and re-read fresh (bounded recovery) instead of sticking on base text.
 */
import { describe, expect, it } from 'vitest'
import { awaitSettled } from '../src/xlsx/cf/text'

function deferredNotify() {
  const listeners = new Set<() => void>()
  return {
    subscribe: (notify: () => void) => {
      listeners.add(notify)
      return () => { listeners.delete(notify) }
    },
    fire: () => { for (const listener of [...listeners]) listener() },
    count: () => listeners.size,
  }
}

describe('awaitSettled resolves on settle, abort, or timeout', () => {
  it('times out when nothing settles and cleans up', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    const result = await awaitSettled(() => false, channel.subscribe, controller.signal, 50)
    expect(result).toBe('timeout')
    expect(channel.count()).toBe(0)
  })

  it('resolves settled when already settled', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    const result = await awaitSettled(() => true, channel.subscribe, controller.signal, 50)
    expect(result).toBe('settled')
    expect(channel.count()).toBe(0)
  })

  it('resolves settled when a notification settles it', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    let settled = false
    const pending = awaitSettled(() => settled, channel.subscribe, controller.signal, 5000)
    expect(channel.count()).toBe(1)
    settled = true
    channel.fire()
    expect(await pending).toBe('settled')
    expect(channel.count()).toBe(0)
  })

  it('resolves settled on abort without waiting for the timeout', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    const pending = awaitSettled(() => false, channel.subscribe, controller.signal, 5000)
    controller.abort()
    expect(await pending).toBe('settled')
    expect(channel.count()).toBe(0)
  })

  it('pre-aborted signals settle immediately', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    controller.abort()
    const result = await awaitSettled(() => false, channel.subscribe, controller.signal, 50)
    expect(result).toBe('settled')
    expect(channel.count()).toBe(0)
  })

  it('balances abort listeners across repeated waits', async () => {
    const channel = deferredNotify()
    const controller = new AbortController()
    let added = 0
    let removed = 0
    const rawAdd = controller.signal.addEventListener.bind(controller.signal)
    const rawRemove = controller.signal.removeEventListener.bind(controller.signal)
    controller.signal.addEventListener = ((...args: [string, EventListener]) => { added++; return (rawAdd as (...a: unknown[]) => void)(...args) }) as typeof controller.signal.addEventListener
    controller.signal.removeEventListener = ((...args: [string, EventListener]) => { removed++; return (rawRemove as (...a: unknown[]) => void)(...args) }) as typeof controller.signal.removeEventListener
    for (let i = 0; i < 20; i++) {
      // eslint-disable-next-line no-await-in-loop
      expect(await awaitSettled(() => false, channel.subscribe, controller.signal, 0)).toBe('timeout')
    }
    expect(added).toBe(20)
    expect(removed).toBe(20)
    expect(channel.count()).toBe(0)
  })

  it('a synchronous notification settles without throwing', async () => {
    let ready = false
    let unsubscribed = false
    const subscribe = (notify: () => void): (() => void) => {
      // Valid eager subscription: marks the predicate ready and notifies
      // before returning the unsubscribe function.
      ready = true
      notify()
      return () => { unsubscribed = true }
    }
    const controller = new AbortController()
    const result = await awaitSettled(() => ready, subscribe, controller.signal, 5000)
    expect(result).toBe('settled')
    expect(unsubscribed).toBe(true)
  })
})
