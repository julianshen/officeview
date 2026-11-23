/** The per-IP session gate (§14: 30 new sessions/min, 8 concurrent connects) and its TEST-only override, which the
 * Playwright WS harness (server/test/app.ts) passes: one browser runs the whole stream project from 127.0.0.1, and
 * once the suite opens more than 30 sessions in a minute the next spec's handshake is refused 429 (Stage 4b: the
 * five `4b default client` specs made `stream-text.spec.ts` "a11y=off" the 31st). */
import { afterAll, describe, expect, it } from 'bun:test'
import { HARNESS_IP_RATES } from './app-options'
import { Client, disposePool, start } from './harness'

afterAll(async () => { await disposePool() })

/** Opens and closes `n` sockets one after another on a frozen clock; how many handshakes were admitted. */
async function admitted(h: Awaited<ReturnType<typeof start>>, n: number): Promise<number> {
  let ok = 0
  for (let i = 0; i < n; i++) {
    const c = new Client(h.url)
    const open = await c.opened().then(() => true, () => false)
    if (open) { ok++; c.ws.close(); await c.closed }
    // The server's close handler releases the concurrent slot: wait for it, so only the per-minute window counts.
    for (let t = 0; h.stream.stats().sockets > 0 && t < 100; t++) await Bun.sleep(10)
  }
  return ok
}

describe('per-IP session gate', () => {
  it('the product default admits 30 new sessions in a minute and refuses the 31st', async () => {
    const h = await start({ now: () => 1_000_000 })
    try { expect(await admitted(h, 31)).toBe(30) } finally { await h.stop() }
  })
  it('`ipRates.perMinute` (TEST-only) replaces the per-minute limit', async () => {
    const h = await start({ now: () => 1_000_000, ipRates: { perMinute: 40 } })
    try { expect(await admitted(h, 41)).toBe(40) } finally { await h.stop() }
  })
  it("the WS harness's override admits more than the product's 30 sessions in a minute", async () => {
    expect(HARNESS_IP_RATES.perMinute).toBeGreaterThan(30)
    const h = await start({ now: () => 1_000_000, ipRates: HARNESS_IP_RATES })
    try { expect(await admitted(h, 61)).toBe(61) } finally { await h.stop() }
  })
})
