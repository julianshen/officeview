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
})
