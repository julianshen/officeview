/**
 * S13 review round 1: the manual Safari/VoiceOver harnesses must tear down on
 * OS signals (scripts/harness/cleanup.ts), the loopback fixture server must
 * answer a malformed path with 400 instead of crashing, and a safaridriver
 * spawn failure must reach the caller as an error, not an unhandled event.
 */
import { describe, expect, it, vi } from 'vitest'
import { request } from 'node:http'
import { installTeardown, SIGNAL_EXIT_CODES, type SignalHost, type TeardownSignal } from '../scripts/harness/cleanup'
import { startFixtureServer } from '../scripts/harness/fixture-server'
import { SafariDriver } from '../scripts/harness/webdriver'

function fakeHost() {
  const listeners = new Map<TeardownSignal, Set<() => void>>()
  const exits: number[] = []
  const host: SignalHost = {
    once: (signal, listener) => { (listeners.get(signal) ?? listeners.set(signal, new Set()).get(signal)!).add(listener) },
    off: (signal, listener) => { listeners.get(signal)?.delete(listener) },
    exit: code => { exits.push(code) },
  }
  const raise = (signal: TeardownSignal) => { for (const listener of [...listeners.get(signal) ?? []]) listener() }
  const count = () => [...listeners.values()].reduce((n, set) => n + set.size, 0)
  return { host, raise, exits, count }
}
const flush = () => new Promise(resolve => setTimeout(resolve, 0))

describe('installTeardown', () => {
  it.each(['SIGINT', 'SIGTERM', 'SIGHUP'] as const)('%s runs the teardown steps in order, then exits with 128 + signal number', async signal => {
    const { host, raise, exits, count } = fakeHost()
    const order: string[] = []
    installTeardown(async () => { order.push('voiceover off'); await flush(); order.push('driver quit'); order.push('servers closed') }, host, () => undefined)
    expect(count()).toBe(3)
    raise(signal)
    expect(exits).toEqual([]) // never exits before the teardown has finished
    await flush(); await flush()
    expect(order).toEqual(['voiceover off', 'driver quit', 'servers closed'])
    expect(exits).toEqual([SIGNAL_EXIT_CODES[signal]])
    expect(count()).toBe(0)
  })

  it('runs at most once across the in-band path and a signal', async () => {
    const { host, raise, exits } = fakeHost()
    const teardown = vi.fn(async () => { await flush() })
    const handle = installTeardown(teardown, host, () => undefined)
    const inBand = handle.run()
    raise('SIGINT')
    await inBand
    await handle.run()
    await flush()
    expect(teardown).toHaveBeenCalledTimes(1)
    expect(exits).toEqual([130])
  })

  it('a failing teardown is logged and the signal still exits', async () => {
    const { host, raise, exits } = fakeHost()
    const log = vi.fn()
    installTeardown(async () => { throw new Error('pkill failed') }, host, log)
    raise('SIGTERM')
    await flush(); await flush()
    expect(log).toHaveBeenCalledWith('teardown failed: pkill failed')
    expect(exits).toEqual([143])
  })

  it('dispose removes the handlers after an in-band teardown', () => {
    const { host, count } = fakeHost()
    installTeardown(async () => undefined, host, () => undefined).dispose()
    expect(count()).toBe(0)
  })
})

describe('fixture server', () => {
  it('answers a malformed path with 400 and keeps serving', async () => {
    const server = await startFixtureServer(process.cwd(), 0)
    try {
      const get = (path: string) => new Promise<number>((resolve, reject) => {
        const req = request({ host: '127.0.0.1', port: server.port, path }, res => { res.resume(); resolve(res.statusCode ?? 0) })
        req.on('error', reject)
        req.end()
      })
      expect(await get('/%E0%A4%A')).toBe(400)
      expect(await get('/not-a-fixture.xlsx')).toBe(404)
      expect(await get('/f1-small.xlsx')).toBe(200)
    } finally {
      await server.close()
    }
  })
})

describe('SafariDriver.start', () => {
  it('a missing safaridriver binary rejects with a clean message instead of an unhandled error event', async () => {
    const driver = new SafariDriver(1, '/nonexistent/safaridriver-for-test')
    if (process.platform !== 'darwin') {
      await expect(driver.start(2000)).rejects.toThrow(/needs macOS/)
      return
    }
    await expect(driver.start(2000)).rejects.toThrow(/safaridriver failed to start \(\/nonexistent\/safaridriver-for-test\): .*ENOENT/)
  })
})
