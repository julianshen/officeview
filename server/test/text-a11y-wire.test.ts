/** Stage 3a Task 8 — the a11y text service over the wire: modes, quota, audit, chunking. */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { AdmissionRefused, TextTooLarge } from '../src/pool'
import { a11yArgs, chunkA11y, COPY_MAX_BYTES, runA11y, runCopy, TEXT_MESSAGE_BYTES, type A11yContext, type TextSlot } from '../src/text/ops'
import type { TextJob, TextReply } from '../src/text/served'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

const NAMES = ['Alfa', 'Bravo', 'Charlie', 'Delta', 'Echo']
const para = (name: string) => `Item ${name} sits here.`
const paragraphsDoc = () => buildDocx(NAMES.map((n) => ({ runs: [{ text: para(n) }] })))
type Policy = AuthDecision['policy']
const capped = (maxParagraphs: number, maxSessionParagraphs?: number): Policy => ({ copy: { maxParagraphs, ...(maxSessionParagraphs !== undefined ? { maxSessionParagraphs } : {}) } })

async function session(o: Parameters<typeof start>[0] = {}) {
  const store = AuditStore.open(':memory:')
  const h = await start({ audit: new AuditTable({ store }), ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, store, stop: () => { c.close(); h.stop() } }
}

/** Sends a text op and collects its reqId's replies until one carries `last: true` or is not of `type`. */
async function collect(c: Client, msg: { reqId: number; type: string } & Record<string, unknown>, type: string) {
  c.send({ token: c.manifest!.session.token, ...msg })
  const parts: Array<{ text: Record<string, any>; bytes: number; raw: string }> = []
  for (;;) {
    const m = await c.next((m) => 'text' in m && m.text.reqId === msg.reqId) as { text: Record<string, any>; bytes: number }
    if (m.text.type !== type) return { parts, end: m.text }
    parts.push({ ...m, raw: JSON.stringify(m.text) })
    if (m.text.last === true) return { parts, end: m.text }
  }
}
type Para = { rect: number[]; text?: string }
/** One a11y request: its paragraphs (concatenated by seq), its raw messages, or its refusal. */
async function a11y(c: Client, reqId: number, unit: number, paragraph?: number) {
  const r = await collect(c, { reqId, type: 'a11y', unit, ...(paragraph !== undefined ? { paragraph } : {}) }, 'a11y-result')
  if (r.end.type === 'refused') return { refused: r.end, paragraphs: [] as Para[], raw: [] as string[] }
  return { paragraphs: r.parts.flatMap((p) => p.text.paragraphs as Para[]), parts: r.parts, raw: r.parts.map((p) => p.raw) }
}
const copyAll = (c: Client, reqId: number, range: object) => collect(c, { reqId, type: 'copy', range }, 'copy-result')
const range = (unit: number, x0: number, y0: number, x1: number, y1: number) => ({ anchor: { unit, x: x0, y: y0 }, focus: { unit, x: x1, y: y1 } })
/** A copy range over paragraphs `from..to`, from their a11y rects. */
const over = (ps: Para[], from: number, to: number) => {
  const a = ps[from]!.rect, b = ps[to]!.rect
  return range(0, a[0]! + 0.5, a[1]! + a[3]! / 2, b[0]! + b[2]! - 0.5, b[1]! + b[3]! / 2)
}

describe('a11y over the wire', () => {
  test("'unit': every paragraph's text, in order; one audit record kind a11y; not counted as copy", async () => {
    const { c, m, h, store, stop } = await session({ doc: await paragraphsDoc(), policy: { copy: 'full' } })
    expect(m.caps.a11y).toBe('unit')
    expect(m.can.textA11y).toBe(true)
    const r = await a11y(c, 10, 0)
    expect(r.parts!.map((p) => [p.text.seq, p.text.last])).toEqual([[0, true]])
    expect(r.paragraphs.map((p) => p.text)).toEqual(NAMES.map(para))
    for (const p of r.paragraphs) { expect(p.rect.length).toBe(4); expect(p.rect[2]).toBeGreaterThan(0); expect(p.rect[3]).toBeGreaterThan(0) }
    // Top to bottom.
    expect(r.paragraphs.map((p) => p.rect[1])).toEqual([...r.paragraphs.map((p) => p.rect[1]!)].sort((a, b) => a - b))
    // A second request is audited again.
    await a11y(c, 11, 0)
    const sid = m.session.sid
    expect(store.copies(sid)).toEqual([
      { sid, kind: 'a11y', paragraphs: 5, fromUnit: 0, toUnit: 0, at: expect.any(Number) },
      { sid, kind: 'a11y', paragraphs: 5, fromUnit: 0, toUnit: 0, at: expect.any(Number) },
    ])
    expect(h.stream.audit.get(sid)!.copied).toEqual({ events: 0, paragraphs: 0 })
    // A paragraph index is not part of 'unit' mode.
    expect((await a11y(c, 12, 0, 1)).refused).toEqual({ reqId: 12, type: 'refused', code: 'policy-refused' })
    expect(store.copies(sid).length).toBe(2)
    stop()
  }, 60_000)

  test("'unit' with copy off (textA11y: true)", async () => {
    const { c, m, stop } = await session({ doc: await paragraphsDoc(), policy: { textA11y: true } })
    expect([m.can.copy, m.caps.a11y]).toEqual([false, 'unit'])
    expect((await a11y(c, 10, 0)).paragraphs.map((p) => p.text)).toEqual(NAMES.map(para))
    stop()
  }, 60_000)

  test("'off' is refused text-unavailable (default policy; textA11y: false with copy full)", async () => {
    for (const policy of [{}, { copy: 'full', textA11y: false }] as Policy[]) {
      const { c, m, store, stop } = await session({ doc: await paragraphsDoc(), policy })
      expect([m.can.textA11y, m.caps.a11y]).toEqual([false, 'off'])
      expect((await a11y(c, 10, 0)).refused).toEqual({ reqId: 10, type: 'refused', code: 'text-unavailable' })
      expect((await a11y(c, 11, 0, 0)).refused).toEqual({ reqId: 11, type: 'refused', code: 'text-unavailable' })
      expect(store.copies(m.session.sid)).toEqual([])
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      stop()
    }
  }, 60_000)

  test("'focused' without paragraph: rects only (no text field anywhere), uncounted, unaudited", async () => {
    const { c, m, store, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2, 3) })
    expect(m.caps.a11y).toBe('focused')
    const r = await a11y(c, 10, 0)
    expect(r.paragraphs.length).toBe(5)
    for (const p of r.paragraphs) { expect(Object.keys(p)).toEqual(['rect']); expect(p.rect.length).toBe(4) }
    for (const raw of r.raw) { expect(raw).not.toContain('"text"'); for (const n of NAMES) expect(raw).not.toContain(n) }
    expect(store.copies(m.session.sid)).toEqual([])
    // Rects-only requests never touch the quota: three paragraphs can still be read.
    for (let i = 0; i < 4; i++) await a11y(c, 20 + i, 0)
    for (let k = 0; k < 3; k++) expect((await a11y(c, 30 + k, 0, k)).paragraphs.filter((p) => p.text !== undefined).map((p) => p.text)).toEqual([para(NAMES[k]!)])
    stop()
  }, 60_000)

  test("'focused' with paragraph k: only that text; counts 1 against the session quota shared with copy; audited", async () => {
    const { c, m, h, store, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2, 3) })
    const sid = m.session.sid
    const rects = (await a11y(c, 10, 0)).paragraphs
    const one = await a11y(c, 11, 0, 1)
    expect(one.paragraphs.length).toBe(5)
    expect(one.paragraphs.map((p) => p.text)).toEqual([undefined, para('Bravo'), undefined, undefined, undefined])
    for (const n of ['Alfa', 'Charlie', 'Delta', 'Echo']) expect(one.raw.join('')).not.toContain(n)
    expect(store.copies(sid)).toEqual([{ sid, kind: 'a11y', paragraphs: 1, fromUnit: 0, toUnit: 0, at: expect.any(Number) }])
    // An out-of-range index is refused and counts nothing.
    expect((await a11y(c, 12, 0, 5)).refused).toEqual({ reqId: 12, type: 'refused', code: 'policy-refused' })
    // A copy of two paragraphs brings the session to 3 of 3.
    const cp = await copyAll(c, 13, over(rects, 2, 3))
    expect(cp.parts.map((p) => p.text.text).join('')).toBe([para('Charlie'), para('Delta')].join('\n'))
    // Now both a11y and copy are exhausted.
    expect((await a11y(c, 14, 0, 4)).refused).toEqual({ reqId: 14, type: 'refused', code: 'copy-quota-exhausted', limit: 3, requested: 1 })
    expect((await copyAll(c, 15, over(rects, 4, 4))).end).toEqual({ reqId: 15, type: 'refused', code: 'copy-quota-exhausted', limit: 3, requested: 1 })
    // Re-reading a paragraph counts again (the client caches; the server does not deduplicate).
    expect((await a11y(c, 16, 0, 1)).refused).toMatchObject({ code: 'copy-quota-exhausted' })
    // Rects are still answered.
    expect((await a11y(c, 17, 0)).paragraphs.length).toBe(5)
    expect(store.copies(sid).map((r) => [r.kind, r.paragraphs])).toEqual([['a11y', 1], ['copy', 2]])
    expect(h.stream.audit.get(sid)!.copied).toEqual({ events: 1, paragraphs: 2 })
    stop()
  }, 60_000)

  test("'focused' with no session cap: every paragraph can be read; the per-copy cap does not apply", async () => {
    const { c, stop } = await session({ doc: await paragraphsDoc(), policy: capped(1) })
    for (let k = 0; k < 5; k++) expect((await a11y(c, 10 + k, 0, k)).paragraphs[k]!.text).toBe(para(NAMES[k]!))
    stop()
  }, 60_000)

  test('malformed arguments are refused policy-refused', async () => {
    const { c, m, store, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2) })
    const bad: Array<Record<string, unknown>> = [{}, { unit: -1 }, { unit: m.units.count }, { unit: 0.5 }, { unit: '0' }, { unit: 0, paragraph: -1 }, { unit: 0, paragraph: 1.5 }, { unit: 0, paragraph: '1' }, { unit: 0, paragraph: null }]
    for (const [i, args] of bad.entries()) {
      expect((await collect(c, { reqId: 40 + i, type: 'a11y', ...args }, 'a11y-result')).end, JSON.stringify(args)).toEqual({ reqId: 40 + i, type: 'refused', code: 'policy-refused' })
    }
    expect(store.copies(m.session.sid)).toEqual([])
    stop()
  }, 60_000)

  test('an empty unit answers one empty final message and writes no record', async () => {
    const { c, m, store, stop } = await session({ doc: await buildDocx([{ runs: [{ text: '' }] }]), policy: { copy: 'full' } })
    const r = await a11y(c, 10, 0)
    expect(r.parts!.map((p) => p.text)).toEqual([{ reqId: 10, type: 'a11y-result', seq: 0, paragraphs: [], last: true }])
    expect(store.copies(m.session.sid)).toEqual([])
    stop()
  }, 60_000)

  test('R5: a11y is charged the text rate (refusal only)', async () => {
    let t = Date.now()
    const { c, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2), now: () => t })
    for (let i = 0; i < 59; i++) { t += 200; expect((await a11y(c, 100 + i, 0)).paragraphs.length).toBe(5) }
    t += 200; expect((await a11y(c, 10, 0, 0)).paragraphs[0]!.text).toBe(para('Alfa'))   // the 60th text op
    t += 200; expect((await a11y(c, 11, 0)).refused).toEqual({ reqId: 11, type: 'refused', code: 'rate-limited' })
    t += 200; expect((await a11y(c, 12, 0, 1)).refused).toEqual({ reqId: 12, type: 'refused', code: 'rate-limited' })
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    stop()
  }, 120_000)
})

describe('T2 over the wire: marker never in any a11y answer for the 13 constructs', async () => {
  const fixtures = await augmentedFixtures()
  test('13 fixtures', () => expect(fixtures.length).toBe(13))
  const leaks = (raw: string[], marker: string) => raw.some((r) => [marker, 'MARKER', marker.slice(0, 8), marker.toLowerCase()].some((q) => r.includes(q)))
  for (const f of fixtures) test(f.id, async () => {
    // 'unit' mode: every unit's whole text.
    const u = await session({ doc: f.withHidden, policy: { contentPolicy: 'visible-only', copy: 'full' } })
    let reqId = 10, served = ''
    for (let unit = 0; unit < u.m.units.count; unit++) {
      const r = await a11y(u.c, reqId++, unit)
      expect(r.refused, `${f.id} unit ${unit}`).toBeUndefined()
      expect(leaks(r.raw, f.marker), `${f.id} unit ${unit}`).toBe(false)
      served += r.paragraphs.map((p) => p.text).join('\n')
    }
    // Non-vacuous: the visible control is served (the generated xlsx crops it off-surface: withheld).
    if (f.kind !== 'xlsx') expect(served, `${f.id} control`).toContain(f.control)
    u.stop()
    // 'focused' mode (no session cap): rects only, then every paragraph by index.
    const x = await session({ doc: f.withHidden, policy: { contentPolicy: 'visible-only', ...capped(1) } })
    let read = ''
    for (let unit = 0; unit < x.m.units.count; unit++) {
      const r = await a11y(x.c, reqId++, unit)
      expect(r.refused, `${f.id} focused unit ${unit}`).toBeUndefined()
      expect(leaks(r.raw, f.marker)).toBe(false)
      for (let k = 0; k < r.paragraphs.length; k++) {
        const one = await a11y(x.c, reqId++, unit, k)
        expect(one.refused).toBeUndefined()
        expect(leaks(one.raw, f.marker), `${f.id} unit ${unit} paragraph ${k}`).toBe(false)
        read += `${one.paragraphs[k]!.text}\n`
      }
    }
    if (f.kind !== 'xlsx') expect(read, `${f.id} focused control`).toContain(f.control)
    x.stop()
  }, 60_000)
})

describe('ops.ts a11y with a fake pool', () => {
  const slot = (): TextSlot => ({ copiedParagraphs: 0 })
  const reply = (...ps: Para[]): TextReply => ({ op: 'a11y', paragraphs: ps })
  function ctx(o: Partial<A11yContext> & Pick<A11yContext, 'job'>) {
    const sent: Array<Record<string, any>> = [], refused: Array<[string, object | undefined]> = [], records: object[] = []
    const c: A11yContext = { slot: slot(), reqId: 7, unit: 2, mode: () => 'unit', caps: () => ({ maxParagraphs: null, maxSessionParagraphs: null }), record: (r) => { records.push(r) }, alive: () => true, send: (m) => sent.push(m as never), refuse: (code, nums) => refused.push([code, nums]), ...o }
    return { c, sent, refused, records }
  }

  test("the job names the unit (and the paragraph); 'unit' records the count", async () => {
    const jobs: TextJob[] = []
    const x = ctx({ job: async (j) => { jobs.push(j); return reply({ rect: [1, 2, 3, 4], text: 'a' }, { rect: [5, 6, 7, 8], text: 'b' }) } })
    await runA11y(x.c)
    expect(jobs).toEqual([{ op: 'a11y', unit: 2 }])
    expect(x.records).toEqual([{ paragraphs: 2 }])
    expect(x.sent).toEqual([{ reqId: 7, type: 'a11y-result', seq: 0, paragraphs: [{ rect: [1, 2, 3, 4], text: 'a' }, { rect: [5, 6, 7, 8], text: 'b' }], last: true }])
    expect(x.c.slot.copiedParagraphs).toBe(0)
    const f = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 1, maxSessionParagraphs: null }), paragraph: 1, job: async (j) => { jobs.push(j); return reply({ rect: [1, 2, 3, 4] }, { rect: [5, 6, 7, 8], text: 'b' }) } })
    await runA11y(f.c)
    expect(jobs.at(-1)).toEqual({ op: 'a11y', unit: 2, paragraph: 1 })
    expect(f.records).toEqual([{ paragraphs: 1 }])
    expect(f.c.slot.copiedParagraphs).toBe(1)
  })

  test("'focused' rects-only strips text even when the worker sends it; extra fields never ride along", async () => {
    const x = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: 3 }), job: async () => reply({ rect: [1, 2, 3, 4], text: 'secret', extra: 1 } as Para) })
    await runA11y(x.c)
    expect(x.sent).toEqual([{ reqId: 7, type: 'a11y-result', seq: 0, paragraphs: [{ rect: [1, 2, 3, 4] }], last: true }])
    expect([x.records, x.c.slot.copiedParagraphs]).toEqual([[], 0])
    // With k, only paragraph k keeps its text, whatever the worker sent.
    const k = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: 3 }), paragraph: 0, job: async () => reply({ rect: [1, 2, 3, 4], text: 'mine' }, { rect: [1, 2, 3, 4], text: 'other' }) })
    await runA11y(k.c)
    expect(k.sent[0]!.paragraphs).toEqual([{ rect: [1, 2, 3, 4], text: 'mine' }, { rect: [1, 2, 3, 4] }])
  })

  test('a paragraph with no rects gets a zero rect (indices never shift)', async () => {
    const x = ctx({ job: async () => reply({ rect: [], text: 'a' }, { rect: [1, 2, 3, 4], text: 'b' }) })
    await runA11y(x.c)
    expect(x.sent[0]!.paragraphs).toEqual([{ rect: [0, 0, 0, 0], text: 'a' }, { rect: [1, 2, 3, 4], text: 'b' }])
  })

  test('mode is read when the turn starts: off ⇒ text-unavailable; a paragraph in unit mode ⇒ policy-refused; out of range ⇒ policy-refused, uncounted', async () => {
    const off = ctx({ mode: () => 'off', job: async () => { throw new Error('unreachable') } })
    await runA11y(off.c)
    const para = ctx({ paragraph: 0, job: async () => { throw new Error('unreachable') } })
    await runA11y(para.c)
    const range = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: null }), paragraph: 3, job: async () => reply({ rect: [1, 2, 3, 4], text: 'a' }) })
    await runA11y(range.c)
    expect([off.refused, para.refused, range.refused]).toEqual([[['text-unavailable', undefined]], [['policy-refused', undefined]], [['policy-refused', undefined]]])
    for (const x of [off, para, range]) expect([x.sent, x.records, x.c.slot.copiedParagraphs]).toEqual([[], [], 0])
  })

  test('the session quota: used + 1 over the cap is refused with limit and requested 1', async () => {
    const x = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: 4 }), paragraph: 0, job: async () => reply({ rect: [1, 2, 3, 4], text: 'a' }) })
    x.c.slot.copiedParagraphs = 4
    await runA11y(x.c)
    expect(x.refused).toEqual([['copy-quota-exhausted', { limit: 4, requested: 1 }]])
    expect([x.sent, x.records, x.c.slot.copiedParagraphs]).toEqual([[], [], 4])
  })

  test('a failed audit write refuses and counts nothing', async () => {
    const x = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: 4 }), paragraph: 0, record: () => { throw new Error('disk') }, job: async () => reply({ rect: [1, 2, 3, 4], text: 'a' }) })
    await runA11y(x.c)
    expect(x.refused).toEqual([['policy-refused', undefined]])
    expect([x.sent, x.c.slot.copiedParagraphs]).toEqual([[], 0])
  })

  test('over the client limits (4 MiB of text, 20,000 paragraphs) is refused copy-too-large with the 4 MiB limit', async () => {
    const big = ctx({ job: async () => reply(...Array.from({ length: 100 }, () => ({ rect: [1, 2, 3, 4], text: 'x'.repeat(42_000) }))) })
    await runA11y(big.c)
    expect(big.refused).toEqual([['copy-too-large', { limit: COPY_MAX_BYTES, requested: 4_200_000 }]])
    const many = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: null }), job: async () => reply(...Array.from({ length: 20_001 }, () => ({ rect: [1, 2, 3, 4] }))) })
    await runA11y(many.c)
    expect(many.refused).toEqual([['copy-too-large', { limit: COPY_MAX_BYTES }]])
    const huge = ctx({ job: async () => { throw new TextTooLarge([2]) } })
    await runA11y(huge.c)
    expect(huge.refused).toEqual([['copy-too-large', { limit: COPY_MAX_BYTES }]])
    // One paragraph too large for one message cannot be sent in the fixed wire shape.
    const long = ctx({ job: async () => reply({ rect: [1, 2, 3, 4], text: 'y'.repeat(TEXT_MESSAGE_BYTES) }) })
    await runA11y(long.c)
    expect(long.refused).toEqual([['copy-too-large', { limit: TEXT_MESSAGE_BYTES }]])
    for (const x of [big, many, huge, long]) expect([x.sent, x.records]).toEqual([[], []])
    // Exactly 20,000 rects-only paragraphs are sent.
    const edge = ctx({ mode: () => 'focused', caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: null }), job: async () => reply(...Array.from({ length: 20_000 }, () => ({ rect: [1, 2, 3, 4] }))) })
    await runA11y(edge.c)
    expect(edge.refused).toEqual([])
    expect(edge.sent.reduce((n, m) => n + m.paragraphs.length, 0)).toBe(20_000)
  })

  test('pool errors: AdmissionRefused is rate-limited, others policy-refused', async () => {
    const out: Array<[string, object | undefined]> = []
    for (const job of [async () => { throw new AdmissionRefused('session-queue') }, async () => { throw new Error('dead') }, async () => ({ missing: [1] }) as TextReply, async () => ({ op: 'copy', text: '', paragraphs: 0 }) as TextReply]) {
      const x = ctx({ job }); await runA11y(x.c); out.push(...x.refused)
    }
    expect(out).toEqual([['rate-limited', undefined], ['policy-refused', undefined], ['policy-refused', undefined], ['policy-refused', undefined]])
  })

  test('a11y and copy share one queue per sid: the shared counter is never raced', async () => {
    const s = slot(), order: string[] = []
    let release!: () => void
    const gate = new Promise<void>((r) => { release = r })
    const caps = () => ({ maxParagraphs: 2, maxSessionParagraphs: 2 })
    const first = ctx({ mode: () => 'focused', caps, paragraph: 0, slot: s, job: async () => { await gate; order.push('a11y'); return reply({ rect: [1, 2, 3, 4], text: 'a' }) } })
    const sent: unknown[] = [], refused: unknown[] = []
    const copy = runCopy({ slot: s, reqId: 8, range: { all: true }, unitCount: 1, caps, record: () => {}, alive: () => true, send: (m) => sent.push(m), refuse: (code, nums) => refused.push([code, nums]), job: async () => { order.push('copy'); return { op: 'copy', text: 'x\ny', paragraphs: 2 } } })
    const a = runA11y(first.c)
    void copy
    // The copy was queued first; it runs before the a11y and takes the whole quota.
    release()
    await Promise.all([copy, a])
    expect(order).toEqual(['copy', 'a11y'])
    expect(first.refused).toEqual([['copy-quota-exhausted', { limit: 2, requested: 1 }]])
    expect(s.copiedParagraphs).toBe(2)
    // And an a11y queued first holds the copy back until it is done.
    const t = slot(); const seen: string[] = []
    let open!: () => void
    const g2 = new Promise<void>((r) => { open = r })
    const a2 = ctx({ mode: () => 'focused', caps, paragraph: 0, slot: t, job: async () => { await g2; seen.push('a11y'); return reply({ rect: [1, 2, 3, 4], text: 'a' }) } })
    const p1 = runA11y(a2.c)
    const copyRefused: unknown[] = []
    const p2 = runCopy({ slot: t, reqId: 9, range: { all: true }, unitCount: 1, caps, record: () => {}, alive: () => true, send: () => {}, refuse: (code, nums) => copyRefused.push([code, nums]), job: async () => { seen.push('copy'); return { op: 'copy', text: 'x\ny', paragraphs: 2 } } })
    open()
    await Promise.all([p1, p2])
    expect(seen).toEqual(['a11y', 'copy'])
    expect(copyRefused).toEqual([['copy-quota-exhausted', { limit: 2, requested: 2 }]])
    expect(t.copiedParagraphs).toBe(1)
  })

  test('a11yArgs validates and rebuilds the arguments', () => {
    expect(a11yArgs({ unit: 1, policy: 'x' }, 2)).toEqual({ unit: 1 })
    expect(a11yArgs({ unit: 0, paragraph: 0 }, 2)).toEqual({ unit: 0, paragraph: 0 })
    for (const bad of [{}, { unit: 2 }, { unit: -1 }, { unit: 1.5 }, { unit: '1' }, { unit: 0, paragraph: -1 }, { unit: 0, paragraph: 0.5 }, { unit: 0, paragraph: null }, { unit: 0, paragraph: '0' }, { unit: 0, paragraph: 2 ** 53 }]) {
      expect(a11yArgs(bad, 2), JSON.stringify(bad)).toBeUndefined()
    }
  })

  test('chunkA11y: ≤ limit, seq from 0, last boolean on every chunk, never an empty non-final chunk, lossless', () => {
    const size = (m: unknown) => new TextEncoder().encode(JSON.stringify(m)).length
    const ps = Array.from({ length: 3000 }, (_, i) => ({ rect: [i, i + 0.5, 3, 4], text: `para "${i}" é中😀` }))
    const ms = chunkA11y(123, ps)!
    expect(ms.length).toBeGreaterThan(2)
    for (const m of ms) { expect(size(m)).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES); expect(m.paragraphs.length).toBeGreaterThan(0) }
    expect(ms.map((m) => m.seq)).toEqual(ms.map((_, i) => i))
    expect(ms.map((m) => m.last)).toEqual([...ms.slice(1).map(() => false), true])
    expect(ms.flatMap((m) => m.paragraphs)).toEqual(ps)
    for (const limit of [130, 200, 333]) for (const m of chunkA11y(99999, ps.slice(0, 40), limit)!) expect(size(m)).toBeLessThanOrEqual(limit)
    expect(chunkA11y(5, [])).toEqual([{ reqId: 5, type: 'a11y-result', seq: 0, paragraphs: [], last: true }])
    expect(chunkA11y(5, [{ rect: [1, 2, 3, 4], text: 'z'.repeat(200) }], 150)).toBeUndefined()
  })
})
