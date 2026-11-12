/** Stage 3a Task 6 — the client's text requests (StreamSession.text) and the
 * RemoteOfficeDoc search/select/copy/a11y API, against the real server (Task 8
 * added the server's a11y op; the stub-transport a11y tests remain for answers
 * the server never sends). */
import { afterAll, describe, expect, test } from 'bun:test'
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { openRemoteDocument } from '../../src/stream/remote/open'
import type { RemoteOfficeDoc, RemoteRange, RemoteSearchHit } from '../../src/stream/remote/remote-doc'
import type { StreamClient, StreamNotice } from '../../src/stream/remote/session'
import type { AuthDecision } from '../src/hooks'
import { serverBackend } from '../src/backend'
import { ORIGIN, disposePool, start } from './harness'

afterAll(() => disposePool())
const backend = serverBackend()
type Json = Record<string, any>
type Sock = WebSocket & { terminate?: () => void }

/**
 * A client factory over real sockets with taps: `out` sees every message the
 * session sends (return false to swallow it; `inject` delivers a synthetic
 * server message), `in` every text message it receives (return false to drop).
 */
function tapped(tap: { out?: (m: Json, socket: number, inject: (m: object) => void) => boolean | void; in?: (m: Json, socket: number) => boolean | void } = {}) {
  const sockets: Sock[] = [], sent: Array<Json & { socket: number }> = []
  const client = (s: { url: string }): StreamClient => {
    const ws = new WebSocket(s.url, { headers: { Origin: ORIGIN } } as never) as Sock
    const i = sockets.push(ws) - 1
    const listeners: Array<(e: MessageEvent) => void> = []
    const inject = (m: object) => setTimeout(() => { for (const l of listeners) l({ data: JSON.stringify(m) } as MessageEvent) }, 0)
    return {
      get readyState() { return ws.readyState },
      get binaryType() { return ws.binaryType },
      set binaryType(t: BinaryType) { ws.binaryType = t },
      send(d: string) {
        const m = JSON.parse(d) as Json
        sent.push({ ...m, socket: i })
        if (tap.out?.(m, i, inject) === false) return
        ws.send(d)
      },
      close: (code?: number, reason?: string) => ws.close(code, reason),
      addEventListener(type: string, l: (e: Event) => void) {
        if (type !== 'message') return ws.addEventListener(type, l)
        listeners.push(l as (e: MessageEvent) => void)
        ws.addEventListener('message', (e) => {
          if (typeof e.data === 'string' && tap.in?.(JSON.parse(e.data), i) === false) return
          l(e)
        })
      },
      removeEventListener() {},
    } as unknown as StreamClient
  }
  return { client, sockets, sent }
}

const NAMES = ['Alfa', 'Bravo', 'Charlie', 'Delta', 'Echo']
const para = (name: string) => `Item ${name} sits here.`
const paragraphsDoc = () => buildDocx(NAMES.map((n) => ({ runs: [{ text: para(n) }] })))
const lines = (n: number, line: string) => buildDocx(Array.from({ length: n }, () => ({ runs: [{ text: line }] })))
const capped = (maxParagraphs: number, maxSessionParagraphs?: number): AuthDecision['policy'] => ({ copy: { maxParagraphs, ...(maxSessionParagraphs !== undefined ? { maxSessionParagraphs } : {}) } })
const failure = (p: Promise<unknown>) => p.then(() => { throw new Error('expected a rejection') }, (e) => e as StreamSessionError)

async function open(o: Parameters<typeof start>[0] = {}, tap: Parameters<typeof tapped>[0] = {}) {
  const h = await start(o), t = tapped(tap), notices: StreamNotice[] = []
  const doc = await openRemoteDocument({ kind: 'stream', url: h.url }, { client: t.client, backend, onNotice: (n) => notices.push(n) })
  return { h, t, doc, notices, stop: () => { doc.close(); h.stop() } }
}

/** A range from the start of paragraph `from` to the end of paragraph `to`, from the search rects. */
async function paragraphs(doc: RemoteOfficeDoc, from: number, to: number): Promise<RemoteRange> {
  const box = async (q: string) => {
    let hit: RemoteSearchHit | undefined
    await doc.search(q, { caseSensitive: true, onHits: (h) => { hit = h[0] } })
    const r = hit!.rects
    const x = Math.min(...r.map((q) => q.x)), y = Math.min(...r.map((q) => q.y))
    return { unit: hit!.unit, x, y, w: Math.max(...r.map((q) => q.x + q.width)) - x, h: Math.max(...r.map((q) => q.y + q.height)) - y }
  }
  const a = await box(para(NAMES[from]!)), b = await box(para(NAMES[to]!))
  return { anchor: { unit: a.unit, x: a.x + 0.5, y: a.y + a.h / 2 }, focus: { unit: b.unit, x: b.x + b.w - 0.5, y: b.y + b.h / 2 } }
}

describe('RemoteOfficeDoc text API', () => {
  test('search resolves with its total; hits arrive through onHits as HighlightRects', async () => {
    const { doc, stop } = await open()
    const calls: RemoteSearchHit[][] = []
    const r = await doc.search('subtitle', { onHits: (h) => calls.push(h) })
    expect(r).toEqual({ total: 1, truncated: false })
    expect(calls.length).toBe(1)
    const [hit] = calls[0]!
    expect(hit!.unit).toBe(0)
    expect(hit!.rects.length).toBeGreaterThan(0)
    for (const q of hit!.rects) { expect(Object.keys(q).sort()).toEqual(['height', 'width', 'x', 'y']); expect(q.width).toBeGreaterThan(0) }
    expect(await doc.search('subtitle', { caseSensitive: true })).toEqual({ total: 0, truncated: false })
    stop()
  }, 60_000)

  test('a multi-message search: each onHits call carries every hit so far', async () => {
    const { doc, stop } = await open({ doc: await lines(900, 'alpha beta gamma delta') })
    const sizes: number[] = []
    let last: RemoteSearchHit[] = []
    const r = await doc.search('beta', { onHits: (h) => { sizes.push(h.length); last = h } })
    expect(r).toEqual({ total: 900, truncated: false })
    expect(sizes.length).toBeGreaterThan(1)
    expect(sizes).toEqual([...sizes].sort((a, b) => a - b))
    expect(last.length).toBe(900)
    stop()
  }, 120_000)

  test('copy over maxParagraphs rejects copy-too-large with limit and requested; select is not capped', async () => {
    const { doc, stop } = await open({ doc: await paragraphsDoc(), policy: capped(2) })
    expect(await doc.copy(await paragraphs(doc, 0, 1))).toBe([para('Alfa'), para('Bravo')].join('\n'))
    const three = await paragraphs(doc, 0, 2)
    const e = await failure(doc.copy(three))
    expect(e).toBeInstanceOf(StreamSessionError)
    expect([e.code, e.limit, e.requested]).toEqual(['copy-too-large', 2, 3])
    const s = await doc.select(three)
    expect(s.paragraphs).toBe(3)
    expect(s.units.map((u) => u.unit)).toEqual([0])
    expect(s.units[0]!.rects.length).toBeGreaterThan(0)
    stop()
  }, 60_000)

  test('copy-quota-exhausted carries the session cap', async () => {
    const { doc, stop } = await open({ doc: await paragraphsDoc(), policy: capped(1, 1) })
    const one = await paragraphs(doc, 1, 1)
    expect(await doc.copy(one)).toBe(para('Bravo'))
    const e = await failure(doc.copy(one))
    expect([e.code, e.limit, e.requested]).toEqual(['copy-quota-exhausted', 1, 1])
    stop()
  }, 60_000)

  test('copy off: select, copy and a11y reject text-unavailable locally; the notice is emitted once', async () => {
    const { doc, t, notices, stop } = await open({ doc: await paragraphsDoc() })
    expect(doc.manifest.can.copy).toBe(false)
    expect(doc.manifest.can.textA11y).toBe(false)
    for (const p of [doc.copy({ all: true }), doc.select({ all: true }), doc.a11y(0), doc.copy({ all: true })]) {
      expect((await failure(p)).code).toBe('text-unavailable')
    }
    expect(notices.filter((n) => n.code === 'text-unavailable').length).toBe(1)
    expect(t.sent.filter((m) => ['copy', 'select', 'a11y'].includes(m.type))).toEqual([])
    // Search is always offered.
    expect((await doc.search('Item')).total).toBe(5)
    stop()
  }, 60_000)

  test('a resume during a search re-issues it; hits restart and the search still resolves', async () => {
    let dropped = false
    const { doc, t, notices, stop } = await open({ doc: await lines(900, 'alpha beta gamma delta') }, {
      in: (m, socket) => {
        if (socket !== 0 || (m.type !== 'search-result' && m.type !== 'search-done')) return
        // The first batch is delivered; then the socket drops with the search in
        // flight (a warm index may already have sent the rest: it is lost with it).
        if (dropped) return false
        dropped = true
        setTimeout(() => t.sockets[0]!.terminate!(), 0)
      },
    })
    const calls: RemoteSearchHit[][] = []
    const r = await doc.search('beta', { onHits: (h) => calls.push(h) })
    expect(r).toEqual({ total: 900, truncated: false })
    expect(t.sockets.length).toBe(2)
    expect(notices.some((n) => n.code === 'resumed')).toBe(true)
    const searches = t.sent.filter((m) => m.type === 'search')
    expect(searches.map((m) => m.socket)).toEqual([0, 1])
    expect(searches[1]!.reqId).not.toBe(searches[0]!.reqId)
    // The restart is announced as an empty list; the hits after it are the whole answer.
    const reset = calls.findIndex((h) => h.length === 0)
    expect(reset).toBeGreaterThan(0)
    expect(calls.at(-1)!.length).toBe(900)
    stop()
  }, 120_000)

  test('a resume during a copy rejects it session-rotated and never re-issues it', async () => {
    let dropped = false
    const { doc, h, t, stop } = await open({ doc: await paragraphsDoc(), policy: capped(2) }, {
      in: (m, socket) => {
        if (socket !== 0 || m.type !== 'copy-result') return
        // The server answered (and counted) the copy; the answer is lost with the socket.
        if (!dropped) { dropped = true; setTimeout(() => t.sockets[0]!.terminate!(), 0) }
        return false
      },
    })
    const one = await paragraphs(doc, 0, 0)
    const e = await failure(doc.copy(one))
    expect(e.code).toBe('session-rotated')
    const sid = doc.manifest.session.sid
    expect(h.stream.audit.get(sid)!.copied).toEqual({ events: 1, paragraphs: 1 })
    // After the resume the session works; the lost copy was not sent again.
    expect(await doc.copy(one)).toBe(para('Alfa'))
    expect(t.sent.filter((m) => m.type === 'copy').map((m) => m.socket)).toEqual([0, 1])
    expect(h.stream.audit.get(sid)!.copied).toEqual({ events: 2, paragraphs: 2 })
    stop()
  }, 60_000)

  test('aborting a search rejects it and stops onHits', async () => {
    const { doc, stop } = await open({ doc: await lines(900, 'alpha beta gamma delta') })
    const ac = new AbortController()
    let calls = 0
    const e = await failure(doc.search('beta', { signal: ac.signal, onHits: () => { calls++; ac.abort() } }))
    expect((e as Error).name).toBe('AbortError')
    await Bun.sleep(500)
    expect(calls).toBe(1)
    expect((await doc.search('gamma')).total).toBe(900)
    stop()
  }, 120_000)

  test('a newer select settles the older one (rejected closed); the newer resolves', async () => {
    const { doc, stop } = await open({ doc: await paragraphsDoc(), policy: capped(5) })
    const range = await paragraphs(doc, 0, 1)
    const older = doc.select({ all: true }), newer = doc.select(range)
    const e = await failure(older)
    expect(e.code).toBe('closed')
    expect((await newer).paragraphs).toBe(2)
    stop()
  }, 60_000)

  test('select chunks that repeat a unit are merged into one hit', async () => {
    const { doc, stop } = await open({ doc: await lines(100, 'alpha beta'), policy: capped(5) }, {
      out: (m, _s, inject) => {
        if (m.type !== 'select') return
        inject({ reqId: m.reqId, type: 'select-result', units: [{ unit: 0, rects: [1, 2, 3, 4] }], paragraphs: 2 })
        inject({ reqId: m.reqId, type: 'select-result', units: [{ unit: 0, rects: [5, 6, 7, 8] }, { unit: 1, rects: [0, 0, 1, 1] }], paragraphs: 2, last: true })
        return false
      },
    })
    expect(doc.manifest.units.count).toBeGreaterThan(1)
    const s = await doc.select({ all: true })
    expect(s).toEqual({
      paragraphs: 2,
      units: [
        { unit: 0, rects: [{ x: 1, y: 2, width: 3, height: 4 }, { x: 5, y: 6, width: 7, height: 8 }] },
        { unit: 1, rects: [{ x: 0, y: 0, width: 1, height: 1 }] },
      ],
    })
    stop()
  }, 60_000)

  test('a11y against the server: unit mode reads every paragraph; focused mode reads rects, then one paragraph, against the quota', async () => {
    const unit = await open({ doc: await paragraphsDoc(), policy: { copy: 'full' } })
    expect(unit.doc.manifest.caps.a11y).toBe('unit')
    const all = await unit.doc.a11y(0)
    expect(all.map((p) => p.text)).toEqual(NAMES.map(para))
    for (const p of all) expect(p.rect.width).toBeGreaterThan(0)
    expect((await failure(unit.doc.a11y(0, 1))).code).toBe('policy-refused')
    unit.stop()
    const focused = await open({ doc: await paragraphsDoc(), policy: capped(2, 1) })
    expect(focused.doc.manifest.caps.a11y).toBe('focused')
    const rects = await focused.doc.a11y(0)
    expect(rects.length).toBe(5)
    for (const p of rects) expect('text' in p).toBe(false)
    expect((await focused.doc.a11y(0, 2)).map((p) => p.text)).toEqual([undefined, undefined, para('Charlie'), undefined, undefined])
    const e = await failure(focused.doc.a11y(0, 3))
    expect([e.code, e.limit, e.requested]).toEqual(['copy-quota-exhausted', 1, 1])
    focused.stop()
  }, 60_000)

  test('a11y (stub transport): paragraphs concatenate by seq and resolve on last', async () => {
    const { doc, t, stop } = await open({ doc: await paragraphsDoc(), policy: { textA11y: true } }, {
      out: (m, _s, inject) => {
        if (m.type !== 'a11y') return
        inject({ reqId: m.reqId, type: 'a11y-result', seq: 0, paragraphs: [{ rect: [1, 2, 3, 4], text: 'one' }], last: false })
        inject({ reqId: m.reqId, type: 'a11y-result', seq: 1, paragraphs: [{ rect: [5, 6, 7, 8] }], last: true })
        return false
      },
    })
    expect(doc.manifest.can.textA11y).toBe(true)
    expect(await doc.a11y(0, 3)).toEqual([{ rect: { x: 1, y: 2, width: 3, height: 4 }, text: 'one' }, { rect: { x: 5, y: 6, width: 7, height: 8 } }])
    const sent = t.sent.find((m) => m.type === 'a11y')!
    expect(sent).toMatchObject({ type: 'a11y', unit: 0, paragraph: 3, token: doc.manifest.session.token })
    stop()
  }, 60_000)

  test('a11y (stub transport): a malformed answer rejects; a resume rejects an in-flight a11y session-rotated', async () => {
    let mode: 'bad' | 'drop' = 'bad'
    const { doc, t, stop } = await open({ doc: await paragraphsDoc(), policy: { textA11y: true } }, {
      out: (m, socket, inject) => {
        if (m.type !== 'a11y') return
        if (mode === 'bad') inject({ reqId: m.reqId, type: 'a11y-result', seq: 1, paragraphs: [], last: true })
        else setTimeout(() => t.sockets[socket]!.terminate!(), 0)
        return false
      },
    })
    expect((await failure(doc.a11y(0))).code).toBe('policy-refused')
    mode = 'drop'
    expect((await failure(doc.a11y(0))).code).toBe('session-rotated')
    await Bun.sleep(300)
    expect(t.sockets.length).toBe(2)
    expect(t.sent.filter((m) => m.type === 'a11y').map((m) => m.socket)).toEqual([0, 0])
    stop()
  }, 60_000)

  test('session end rejects pending text requests with the typed code', async () => {
    const { doc, h, stop } = await open({ doc: await lines(900, 'alpha beta gamma delta') })
    const p = doc.search('beta')
    h.stream.revoke(doc.manifest.session.sid)
    expect((await failure(p)).code).toBe('revoked')
    stop()
  }, 120_000)
})

/** A doc whose `type` requests are swallowed and answered with `answer(reqId)` (injected in order). */
async function stubbed(type: string, answer: (reqId: number) => object[], policy: AuthDecision['policy'] = capped(5)) {
  return open({ doc: await paragraphsDoc(), policy }, {
    out: (m, _s, inject) => {
      if (m.type !== type) return
      for (const x of answer(m.reqId)) inject(x)
      return false
    },
  })
}
const refusedPolicy = async (p: Promise<unknown>) => { const e = await failure(p); expect(e).toBeInstanceOf(StreamSessionError); expect(e.code).toBe('policy-refused'); return e }
const quads = (n: number) => Array.from({ length: n }, () => [1, 2, 3, 4]).flat()

describe('text answers are bounded and checked (review round 1)', () => {
  test('copy: 100 chunks of 48,000 chars (4.8 MB > 4 MiB) reject policy-refused; the session survives', async () => {
    const { doc, stop } = await stubbed('copy', (reqId) => Array.from({ length: 100 }, (_, seq) => ({ reqId, type: 'copy-result', seq, text: 'a'.repeat(48_000), last: seq === 99 })))
    const e = await refusedPolicy(doc.copy({ all: true }))
    expect(e.message).toContain('4 MiB')
    expect((await doc.search('Item')).total).toBe(5)
    stop()
  }, 60_000)

  test('one text-service message over 48 KiB of UTF-8 rejects (ASCII length under the cap, bytes over it)', async () => {
    // 30,000 'é' = 30,000 UTF-16 units (under 49,152) but 60,000 bytes of UTF-8.
    const { doc, stop } = await stubbed('copy', (reqId) => [{ reqId, type: 'copy-result', seq: 0, text: 'é'.repeat(30_000), last: true }])
    expect((await refusedPolicy(doc.copy({ all: true }))).message).toContain('48 KiB')
    stop()
  }, 60_000)

  test('search: 5,100 hits (> 5,000) reject before the consumer sees the overflow', async () => {
    const sizes: number[] = []
    const { doc, stop } = await stubbed('search', (reqId) => [
      ...Array.from({ length: 6 }, () => ({ reqId, type: 'search-result', matches: Array.from({ length: 850 }, () => ({ unit: 0, rects: [1, 2, 3, 4] })) })),
      { reqId, type: 'search-done', total: 5100, truncated: false },
    ])
    await refusedPolicy(doc.search('x', { onHits: (h) => sizes.push(h.length) }))
    expect(Math.max(...sizes)).toBeLessThanOrEqual(5000)
    stop()
  }, 60_000)

  test('search: a search-done total that differs from the hits delivered rejects', async () => {
    const { doc, stop } = await stubbed('search', (reqId) => [
      { reqId, type: 'search-result', matches: [{ unit: 0, rects: [1, 2, 3, 4] }, { unit: 0, rects: [5, 6, 7, 8] }] },
      { reqId, type: 'search-done', total: 3, truncated: false },
    ])
    expect((await refusedPolicy(doc.search('x'))).message).toContain('total')
    stop()
  }, 60_000)

  test('select: more than 20,000 rects across chunks rejects', async () => {
    const { doc, stop } = await stubbed('select', (reqId) => Array.from({ length: 11 }, (_, i) => ({ reqId, type: 'select-result', units: [{ unit: 0, rects: quads(2000) }], paragraphs: 1, ...(i === 10 ? { last: true } : {}) })))
    await refusedPolicy(doc.select({ all: true }))
    stop()
  }, 60_000)

  test('a11y: more than 20,000 paragraphs, or more than 4 MiB of text, rejects', async () => {
    let mode: 'count' | 'bytes' = 'count'
    const { doc, stop } = await stubbed('a11y', (reqId) => mode === 'count'
      ? Array.from({ length: 14 }, (_, seq) => ({ reqId, type: 'a11y-result', seq, paragraphs: Array.from({ length: 1500 }, () => ({ rect: [1, 2, 3, 4] })), last: seq === 13 }))
      : Array.from({ length: 90 }, (_, seq) => ({ reqId, type: 'a11y-result', seq, paragraphs: [{ rect: [1, 2, 3, 4], text: 'a'.repeat(48_000) }], last: seq === 89 })), { textA11y: true })
    await refusedPolicy(doc.a11y(0))
    mode = 'bytes'
    await refusedPolicy(doc.a11y(0))
    stop()
  }, 60_000)

  test('a hit on a unit the manifest does not have, or a negative width, rejects before onHits; negative x/y are kept', async () => {
    let rects = [1, 2, -3, 4], unit = 999
    const calls: RemoteSearchHit[][] = []
    const { doc, stop } = await stubbed('search', (reqId) => [
      { reqId, type: 'search-result', matches: [{ unit, rects }] },
      { reqId, type: 'search-done', total: 1, truncated: false },
    ])
    expect(doc.manifest.units.count).toBeLessThan(999)
    await refusedPolicy(doc.search('x', { onHits: (h) => calls.push(h) }))
    unit = 0
    await refusedPolicy(doc.search('x', { onHits: (h) => calls.push(h) }))
    rects = [1, 2, 3, -4]
    await refusedPolicy(doc.search('x', { onHits: (h) => calls.push(h) }))
    expect(calls).toEqual([])
    rects = [-1, -2, 3, 4]
    expect(await doc.search('x', { onHits: (h) => calls.push(h) })).toEqual({ total: 1, truncated: false })
    expect(calls).toEqual([[{ unit: 0, rects: [{ x: -1, y: -2, width: 3, height: 4 }] }]])
    stop()
  }, 60_000)

  test('select and a11y geometry are checked the same way', async () => {
    const { doc, stop } = await stubbed('select', (reqId) => [{ reqId, type: 'select-result', units: [{ unit: 999, rects: [1, 2, 3, 4] }], paragraphs: 1, last: true }])
    await refusedPolicy(doc.select({ all: true }))
    stop()
    const b = await stubbed('a11y', (reqId) => [{ reqId, type: 'a11y-result', seq: 0, paragraphs: [{ rect: [1, 2, -3, 4] }], last: true }], { textA11y: true })
    await refusedPolicy(b.doc.a11y(0))
    b.stop()
  }, 60_000)

  test('a live AbortSignal is released by every settled request (success, refusal, supersession, end)', async () => {
    const ac = new AbortController(), sig = ac.signal
    let added = 0, removed = 0
    const add = sig.addEventListener.bind(sig), remove = sig.removeEventListener.bind(sig)
    sig.addEventListener = ((...a: Parameters<typeof add>) => { if (a[0] === 'abort') added++; return add(...a) }) as typeof sig.addEventListener
    sig.removeEventListener = ((...a: Parameters<typeof remove>) => { if (a[0] === 'abort') removed++; return remove(...a) }) as typeof sig.removeEventListener
    const { doc, stop } = await open({ doc: await paragraphsDoc(), policy: capped(5) })
    for (let i = 0; i < 3; i++) await doc.search('Item', { signal: sig })
    await refusedPolicy(doc.search('x'.repeat(300), { signal: sig }))     // refused by the server (query > 256)
    const older = doc.select({ all: true }, sig), newer = doc.select({ all: true }, sig)
    expect((await failure(older)).code).toBe('closed')
    await newer
    const pending = doc.search('Item', { signal: sig })
    doc.close()
    await failure(pending)
    expect(sig.aborted).toBe(false)
    expect(added).toBe(7)
    expect(removed).toBe(added)
    stop()
  }, 60_000)
})

describe('payload-free chunks and the message cap (review round 2)', () => {
  test('copy: 100,000 empty non-final chunks are refused on the first; a later valid final cannot revive it', async () => {
    const { doc, stop } = await stubbed('copy', (reqId) => [
      ...Array.from({ length: 100_000 }, (_, seq) => ({ reqId, type: 'copy-result', seq, text: '', last: false })),
      { reqId, type: 'copy-result', seq: 100_000, text: '', last: true },
    ])
    expect((await refusedPolicy(doc.copy({ all: true }))).message).toContain('empty non-final')
    stop()
  }, 120_000)

  test('copy: an empty first chunk followed by valid text still refuses (refused on the first message)', async () => {
    const { doc, stop } = await stubbed('copy', (reqId) => [
      { reqId, type: 'copy-result', seq: 0, text: '', last: false },
      { reqId, type: 'copy-result', seq: 1, text: 'ok', last: true },
    ])
    await refusedPolicy(doc.copy({ all: true }))
    await Bun.sleep(100)
    // The session survives the refusal.
    expect((await doc.search('Item')).total).toBe(5)
    stop()
  }, 60_000)

  test('select, a11y and search: payload-free non-final chunks refuse', async () => {
    const a = await stubbed('select', (reqId) => [
      { reqId, type: 'select-result', units: [{ unit: 0, rects: [] }], paragraphs: 1 },
      { reqId, type: 'select-result', units: [{ unit: 0, rects: [1, 2, 3, 4] }], paragraphs: 1, last: true },
    ])
    await refusedPolicy(a.doc.select({ all: true }))
    a.stop()
    const b = await stubbed('a11y', (reqId) => [
      { reqId, type: 'a11y-result', seq: 0, paragraphs: [], last: false },
      { reqId, type: 'a11y-result', seq: 1, paragraphs: [{ rect: [1, 2, 3, 4] }], last: true },
    ], { textA11y: true })
    await refusedPolicy(b.doc.a11y(0))
    b.stop()
    const c = await stubbed('search', (reqId) => [
      { reqId, type: 'search-result', matches: [] },
      { reqId, type: 'search-done', total: 0, truncated: false },
    ])
    await refusedPolicy(c.doc.search('x'))
    c.stop()
  }, 60_000)

  test('the single final empty answers stay valid', async () => {
    const a = await stubbed('copy', (reqId) => [{ reqId, type: 'copy-result', seq: 0, text: '', last: true }])
    expect(await a.doc.copy({ all: true })).toBe('')
    a.stop()
    const b = await stubbed('select', (reqId) => [{ reqId, type: 'select-result', units: [], paragraphs: 0, last: true }])
    expect(await b.doc.select({ all: true })).toEqual({ units: [], paragraphs: 0 })
    b.stop()
    const c = await stubbed('a11y', (reqId) => [{ reqId, type: 'a11y-result', seq: 0, paragraphs: [], last: true }], { textA11y: true })
    expect(await c.doc.a11y(0)).toEqual([])
    c.stop()
    const d = await stubbed('search', (reqId) => [{ reqId, type: 'search-done', total: 0, truncated: false }])
    expect(await d.doc.search('x')).toEqual({ total: 0, truncated: false })
    d.stop()
  }, 60_000)

  test('at most 4,096 messages per request: 4,096 resolve, 4,097 refuse', async () => {
    let n = 4096
    const { doc, stop } = await stubbed('copy', (reqId) => Array.from({ length: n }, (_, seq) => ({ reqId, type: 'copy-result', seq, text: 'a', last: seq === n - 1 })))
    expect((await doc.copy({ all: true })).length).toBe(4096)
    n = 4097
    expect((await refusedPolicy(doc.copy({ all: true }))).message).toContain('4096')
    stop()
  }, 60_000)
})

describe('remote path purity', () => {
  test('src/stream/remote and the remote view (src/components/remote) never reach the local text index or selection modules', () => {
    const roots = ['../../src/stream/remote', '../../src/components/remote'].map((d) => resolve(import.meta.dir, d))
    const seen = new Set<string>(), queue = roots.flatMap((root) => readdirSync(root).filter((f) => /\.tsx?$/.test(f)).map((f) => resolve(root, f)))
    while (queue.length) {
      const file = queue.pop()!
      if (seen.has(file)) continue
      seen.add(file)
      const src = readFileSync(file, 'utf8')
      // Value and type imports alike: the remote path must not even name them.
      for (const [, spec] of src.matchAll(/(?:import|export)[^'"]*?from\s+['"](\.[^'"]+)['"]/g)) {
        const base = resolve(dirname(file), spec!)
        const hit = ['.ts', '.tsx', '/index.ts'].map((x) => base + x).find((p) => { try { readFileSync(p); return true } catch { return false } })
        if (hit) queue.push(hit)
      }
    }
    const reached = [...seen].map((f) => f.slice(resolve(import.meta.dir, '../..').length + 1))
    expect(reached.filter((f) => f === 'src/core/search.ts' || f === 'src/core/selection.ts')).toEqual([])
    expect(reached).toContain('src/stream/remote/session.ts')
    expect(reached).toContain('src/components/remote/RemoteText.tsx')
  })
})
