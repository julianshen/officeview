/** Stage 3a Task 4 — search over the wire (`server/src/text/ops.ts` + ws dispatch). */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { AdmissionRefused, TextTooLarge } from '../src/pool'
import { chunkMatches, MAX_SEARCH_MATCHES, runSearch, SEARCH_BATCH_UNITS, TEXT_MESSAGE_BYTES, type SearchJob, type SearchSlot } from '../src/text/ops'
import type { TextReply } from '../src/text/served'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

const KIB48 = 48 * 1024
/** A docx of `n` one-line paragraphs of `line` (≈ 46 lines a page). */
const lines = (n: number, line: string) => buildDocx(Array.from({ length: n }, () => ({ runs: [{ text: line }] })))
async function session(o: Parameters<typeof start>[0] = {}) {
  const h = await start(o), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, stop: () => { c.close(); h.stop() } }
}
const search = (c: Client, reqId: number, query: string, extra: Record<string, unknown> = {}) => c.text({ reqId, type: 'search', query, ...extra })

describe('search over the wire', () => {
  test('search answers rects then search-done with the total', async () => {
    const { c, m, stop } = await session()
    const { parts, end } = await search(c, 10, 'subtitle')
    expect(end).toEqual({ reqId: 10, type: 'search-done', total: 1, truncated: false })
    const matches = parts.flatMap((p) => p.text.matches as Array<{ unit: number; rects: number[] }>)
    expect(matches.length).toBe(1)
    const { w, h } = m.units.list[0]!
    for (const x of matches) {
      expect(x.unit).toBe(0)
      expect(x.rects.length % 4).toBe(0)
      for (let i = 0; i < x.rects.length; i += 4) {
        const [rx, ry, rw, rh] = x.rects.slice(i, i + 4) as [number, number, number, number]
        expect(rw).toBeGreaterThan(0); expect(rh).toBeGreaterThan(0)
        expect(rx + rw).toBeLessThanOrEqual(w); expect(ry + rh).toBeLessThanOrEqual(h)
      }
    }
    // caseSensitive is honoured; no match ⇒ no search-result at all.
    const none = await search(c, 11, 'subtitle', { caseSensitive: true })
    expect(none.parts).toEqual([])
    expect(none.end).toEqual({ reqId: 11, type: 'search-done', total: 0, truncated: false })
    stop()
  }, 60_000)

  test('a second search supersedes the first', async () => {
    // A cold multi-page document, so the first search is still walking when the second arrives.
    const { c, m, stop } = await session({ doc: await lines(900, 'alpha beta gamma delta') })
    expect(m.units.count).toBeGreaterThan(2 * SEARCH_BATCH_UNITS)
    c.send({ reqId: 20, type: 'search', token: m.session.token, query: 'beta' })
    const second = await search(c, 21, 'gamma')
    const first = c.msgs.filter((x) => 'text' in x && x.text.reqId === 20).map((x) => (x as { text: Record<string, any> }).text)
    const ends = first.filter((x) => x.type !== 'search-result')
    expect(ends).toEqual([{ reqId: 20, type: 'search-done', total: first.slice(0, -1).reduce((n, x) => n + x.matches.length, 0), truncated: true }])
    expect(first.at(-1)!.type).toBe('search-done')            // nothing for the old reqId after its done
    expect(second.end).toEqual({ reqId: 21, type: 'search-done', total: 900, truncated: false })
    await Bun.sleep(200)
    expect(c.msgs.filter((x) => 'text' in x && x.text.reqId === 20).length).toBe(first.length)
    stop()
  }, 120_000)

  test('query of 257 chars is refused; 256 is accepted', async () => {
    const { c, stop } = await session()
    expect((await search(c, 30, 'a'.repeat(257))).end).toEqual({ reqId: 30, type: 'refused', code: 'policy-refused' })
    expect((await search(c, 31, '')).end).toEqual({ reqId: 31, type: 'refused', code: 'policy-refused' })
    expect((await search(c, 32, 'x', { caseSensitive: 'yes' })).end).toEqual({ reqId: 32, type: 'refused', code: 'policy-refused' })
    expect((await search(c, 33, 'a'.repeat(256))).end).toEqual({ reqId: 33, type: 'search-done', total: 0, truncated: false })
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    stop()
  }, 60_000)

  test('61st text op in a minute is refused rate-limited and the socket stays open', async () => {
    let t = Date.now()
    const { c, stop } = await session({ now: () => t })
    // 200 ms apart (injected clock): well inside the message budget (20/s, 120/10 s).
    for (let i = 0; i < 60; i++) {
      t += 200
      expect((await search(c, 100 + i, 'subtitle')).end.type).toBe('search-done')
    }
    // Overflows are refusals only: three of them would close 4008 if they counted as breaches.
    for (let i = 0; i < 4; i++) {
      t += 200
      expect((await search(c, 200 + i, 'subtitle')).end).toEqual({ reqId: 200 + i, type: 'refused', code: 'rate-limited' })
    }
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    t += 200; c.unit(300)
    expect((await c.unitMessages(300)).unit.discriminator).toBe(2)
    // The window slides: a minute after the first op, text ops are admitted again.
    t += 60_000
    expect((await search(c, 301, 'subtitle')).end.type).toBe('search-done')
    stop()
  }, 120_000)

  test('search with a stale jti is refused session-rotated', async () => {
    const h = await start(), a = new Client(h.url); const m = await a.hello(); await a.ready()
    const b = new Client(h.url)
    await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } }); await b.ready()
    expect((await b.text({ reqId: 40, type: 'search', query: 'subtitle', token: m.session.token })).end).toEqual({ reqId: 40, type: 'refused', code: 'session-rotated' })
    expect((await b.text({ reqId: 41, type: 'search', query: 'subtitle', token: 'garbage' })).end).toEqual({ reqId: 41, type: 'refused', code: 'policy-refused' })
    expect((await search(b, 42, 'subtitle')).end).toEqual({ reqId: 42, type: 'search-done', total: 1, truncated: false })
    b.close(); h.stop()
  }, 60_000)

  test('no search-result message exceeds 48 KiB', async () => {
    // Every line matches ten times: more than 5,000 matches, many KiB per batch.
    const { c, stop } = await session({ doc: await lines(700, 'e e e e e e e e e e') })
    const { parts, end } = await search(c, 50, 'e')
    expect(end).toEqual({ reqId: 50, type: 'search-done', total: MAX_SEARCH_MATCHES, truncated: true })
    const bytes = parts.map((p) => p.bytes)
    expect(bytes.reduce((n, b) => n + b, 0)).toBeGreaterThan(2 * KIB48)
    for (const b of bytes) expect(b).toBeLessThanOrEqual(KIB48)
    expect(parts.reduce((n, p) => n + p.text.matches.length, 0)).toBe(MAX_SEARCH_MATCHES)
    // Units arrive in order.
    const units = parts.flatMap((p) => p.text.matches.map((x: { unit: number }) => x.unit))
    expect(units).toEqual([...units].sort((x, y) => x - y))
    stop()
  }, 120_000)
})

describe('T2 over the wire: marker never in any search-result for the 13 constructs', async () => {
  const fixtures = await augmentedFixtures()
  test('13 fixtures', () => expect(fixtures.length).toBe(13))
  for (const f of fixtures) test(f.id, async () => {
    const { c, stop } = await session({ doc: f.withHidden, policy: { contentPolicy: 'visible-only' } })
    // The marker and substrings of it that are not in the control (`VIS-<id>` shares the id).
    const queries = [f.marker, 'MARKER', f.marker.slice(0, 8), 'ARKE', f.marker.toLowerCase()]
    let reqId = 60
    for (const q of queries) {
      const { parts, end } = await search(c, reqId, q)
      expect(parts, `${f.id} "${q}"`).toEqual([])
      expect(end, `${f.id} "${q}"`).toEqual({ reqId, type: 'search-done', total: 0, truncated: false })
      reqId++
    }
    // Non-vacuous: the visible control is found (the generated xlsx crops it off-surface: withheld).
    if (f.kind !== 'xlsx') expect((await search(c, reqId, f.control)).end.total, `${f.id} control`).toBeGreaterThanOrEqual(1)
    stop()
  }, 60_000)
})

describe('ops.ts runSearch with a fake pool', () => {
  const run = async (unitCount: number, batch: (job: SearchJob) => Promise<TextReply>, slot: SearchSlot = {}) => {
    const sent: Array<Record<string, any>> = [], refused: string[] = [], jobs: SearchJob[] = []
    await runSearch({ slot, reqId: 7, query: 'q', caseSensitive: false, unitCount, alive: () => true, send: (m) => sent.push(m as never), refuse: (code) => refused.push(code), batch: (j) => { jobs.push(j); return batch(j) } })
    return { sent, refused, jobs, slot }
  }
  const matches = (n: number, unit: number) => Array.from({ length: n }, () => ({ unit, rects: [1, 2, 3, 4] }))

  test('batches of 8 units in order, asking for one more than may still be sent', async () => {
    const r = await run(20, async (j) => ({ op: 'search', matches: matches(2, j.fromUnit) }))
    expect(r.jobs.map((j) => [j.fromUnit, j.toUnit, j.maxMatches])).toEqual([[0, 7, 5001], [8, 15, 4999], [16, 19, 4997]])
    expect(r.sent.at(-1)).toEqual({ reqId: 7, type: 'search-done', total: 6, truncated: false })
    expect(r.slot.search).toBeUndefined()
  })

  test('exactly 5,000 matches is not truncated; one more is', async () => {
    const exact = await run(16, async (j) => ({ op: 'search', matches: j.fromUnit === 0 ? matches(MAX_SEARCH_MATCHES, 0) : [] }))
    expect(exact.sent.at(-1)).toEqual({ reqId: 7, type: 'search-done', total: 5000, truncated: false })
    const over = await run(16, async (j) => ({ op: 'search', matches: j.fromUnit === 0 ? matches(MAX_SEARCH_MATCHES, 0) : matches(1, 8) }))
    expect(over.sent.at(-1)).toEqual({ reqId: 7, type: 'search-done', total: 5000, truncated: true })
    expect(over.sent.filter((m) => m.type === 'search-result').flatMap((m) => m.matches).length).toBe(5000)
  })

  test('zero units: search-done at once, no pool job', async () => {
    const r = await run(0, async () => { throw new Error('unreachable') })
    expect(r.jobs).toEqual([])
    expect(r.sent).toEqual([{ reqId: 7, type: 'search-done', total: 0, truncated: false }])
  })

  test('pool errors are refusals and no search-done follows', async () => {
    const busy = await run(16, async () => { throw new AdmissionRefused('session-queue') })
    expect([busy.refused, busy.sent]).toEqual([['rate-limited'], []])
    const big = await run(16, async (j) => { if (j.fromUnit === 8) throw new TextTooLarge([9]); return { op: 'search', matches: matches(1, 0) } })
    expect(big.refused).toEqual(['policy-refused'])
    expect(big.sent.map((m) => m.type)).toEqual(['search-result'])
    expect(big.slot.search).toBeUndefined()
  })

  test('a dead owner stops the walk silently', async () => {
    let alive = true
    const sent: unknown[] = [], jobs: SearchJob[] = []
    await runSearch({ slot: {}, reqId: 7, query: 'q', caseSensitive: false, unitCount: 24, alive: () => alive, send: (m) => sent.push(m), refuse: () => sent.push('refused'), batch: async (j) => { jobs.push(j); alive = false; return { op: 'search', matches: matches(1, 0) } } })
    expect(jobs.length).toBe(1)
    expect(sent).toEqual([])
  })

  test('supersession answers the old reqId once and the new run waits for its pool job', async () => {
    const slot: SearchSlot = {}
    let release!: () => void
    const gate = new Promise<void>((r) => { release = r })
    const sent: Array<Record<string, any>> = []
    const order: string[] = []
    const a = runSearch({ slot, reqId: 1, query: 'q', caseSensitive: false, unitCount: 16, alive: () => true, send: (m) => sent.push(m as never), refuse: () => {}, batch: async (j) => { order.push(`a${j.fromUnit}`); if (j.fromUnit === 8) await gate; return { op: 'search', matches: matches(1, j.fromUnit) } } })
    await Bun.sleep(5)
    const b = runSearch({ slot, reqId: 2, query: 'q', caseSensitive: false, unitCount: 8, alive: () => true, send: (m) => sent.push(m as never), refuse: () => {}, batch: async (j) => { order.push(`b${j.fromUnit}`); return { op: 'search', matches: [] } } })
    await Bun.sleep(5)
    expect(sent.at(-1)).toEqual({ reqId: 1, type: 'search-done', total: 1, truncated: true })
    expect(order).toEqual(['a0', 'a8'])           // b has not started while a's job runs
    release(); await Promise.all([a, b])
    expect(order).toEqual(['a0', 'a8', 'b0'])
    expect(sent.filter((m) => m.reqId === 1).map((m) => m.type)).toEqual(['search-result', 'search-done'])
    expect(sent.at(-1)).toEqual({ reqId: 2, type: 'search-done', total: 0, truncated: false })
  })
})

describe('chunkMatches', () => {
  test('every message is ≤ the limit and order is kept', () => {
    const ms = Array.from({ length: 4000 }, (_, i) => ({ unit: i >> 8, rects: [i + 0.123456789, 2.5, 3.75, 14.0625] }))
    const out = chunkMatches(3, ms)
    expect(out.length).toBeGreaterThan(1)
    for (const m of out) expect(new TextEncoder().encode(JSON.stringify(m)).length).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES)
    expect(out.flatMap((m) => m.matches)).toEqual(ms)
  })

  test('a single match over the limit is split across messages by whole rects, same unit', () => {
    const rects = Array.from({ length: 4 * 5000 }, (_, i) => i * 1.000001)
    const out = chunkMatches(3, [{ unit: 2, rects }, { unit: 3, rects: [1, 2, 3, 4] }])
    expect(out.length).toBeGreaterThan(1)
    for (const m of out) expect(new TextEncoder().encode(JSON.stringify(m)).length).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES)
    const pieces = out.flatMap((m) => m.matches)
    for (const p of pieces) expect(p.rects.length % 4).toBe(0)
    expect(pieces.filter((p) => p.unit === 2).flatMap((p) => p.rects)).toEqual(rects)
    expect(pieces.at(-1)).toEqual({ unit: 3, rects: [1, 2, 3, 4] })
  })

  test('nothing to send ⇒ no message', () => expect(chunkMatches(3, [])).toEqual([]))
})
