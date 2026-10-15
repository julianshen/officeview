/**
 * S13 review round 1: the manual Safari/VoiceOver harnesses must tear down on
 * OS signals (scripts/harness/cleanup.ts), the loopback fixture server must
 * answer a malformed path with 400 instead of crashing, and a safaridriver
 * spawn failure must reach the caller as an error, not an unhandled event.
 */
import { describe, expect, it, vi } from 'vitest'
import { createServer, request, type Server } from 'node:http'
import { installTeardown, SIGNAL_EXIT_CODES, type SignalHost, type TeardownSignal } from '../scripts/harness/cleanup'
import { startFixtureServer } from '../scripts/harness/fixture-server'
import { QUIT_TIMEOUT_MS, SafariDriver } from '../scripts/harness/webdriver'

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

  it('a never-resolving step cannot hold the process: the signal exits at the deadline, after the first step ran', async () => {
    const { host, raise, exits } = fakeHost()
    const log = vi.fn()
    const order: string[] = []
    installTeardown(async () => { order.push('voiceover off (verified)'); await new Promise(() => undefined); order.push('never') }, host, log, 30)
    raise('SIGINT')
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(exits).toEqual([]) // still inside the deadline
    expect(order).toEqual(['voiceover off (verified)'])
    await new Promise(resolve => setTimeout(resolve, 60))
    expect(exits).toEqual([130])
    expect(order).toEqual(['voiceover off (verified)'])
    expect(log).toHaveBeenCalledWith('teardown did not finish within 30 ms; continuing without it')
  })

  it('the in-band run also settles at the deadline when a step never resolves', async () => {
    const { host } = fakeHost()
    const handle = installTeardown(() => new Promise<void>(() => undefined), host, () => undefined, 30)
    const started = Date.now()
    await handle.run()
    expect(Date.now() - started).toBeLessThan(1000)
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

describe('bounded teardown steps', () => {
  it('fixture server close does not wait for a stalled client connection', async () => {
    const server = await startFixtureServer(process.cwd(), 0)
    const { connect } = await import('node:net')
    // A client that connects and never finishes its request headers.
    const socket = connect(server.port, '127.0.0.1')
    await new Promise<void>(resolve => socket.once('connect', () => resolve()))
    socket.write('GET /f1-small.xlsx HTTP/1.1\r\nHost: 127.0.0.1\r\n')
    await new Promise(resolve => setTimeout(resolve, 50))
    const started = Date.now()
    await server.close()
    expect(Date.now() - started).toBeLessThan(1000)
    socket.destroy()
  }, 10_000)

  it('SafariDriver.quit abandons a session delete that safaridriver never answers', async () => {
    const wedged: Server = createServer(() => undefined) // accepts, never responds
    await new Promise<void>(resolve => wedged.listen(0, '127.0.0.1', () => resolve()))
    const port = (wedged.address() as { port: number }).port
    try {
      const driver = new SafariDriver(port)
      driver.sessionId = 'wedged-session'
      const started = Date.now()
      await driver.quit()
      const elapsed = Date.now() - started
      expect(elapsed).toBeGreaterThanOrEqual(QUIT_TIMEOUT_MS - 100)
      expect(elapsed).toBeLessThan(QUIT_TIMEOUT_MS + 2000)
      expect(driver.sessionId).toBeNull()
    } finally {
      wedged.closeAllConnections()
      await new Promise(resolve => wedged.close(resolve))
    }
  }, 10_000)
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
