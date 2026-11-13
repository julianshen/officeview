/** Stage 3a Task 5 — select and copy over the wire: caps, session quota, audit. */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildTextIndex } from '../../src/core/search'
import { hitTest, normalizeRange, textForRange } from '../../src/core/selection'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { AdmissionRefused, TextTooLarge } from '../src/pool'
import { DocumentHost } from '../src/render'
import { chunkSelect, chunkText, COPY_MAX_BYTES, rangeArg, runCopy, runSelect, TEXT_MESSAGE_BYTES, type CopyContext, type TextSlot } from '../src/text/ops'
import type { TextJob, TextRange, TextReply, UnitPoint } from '../src/text/served'
import { Client, corpus, disposePool, pool, start } from './harness'

afterAll(() => disposePool())

const KIB48 = 48 * 1024
const NAMES = ['Alfa', 'Bravo', 'Charlie', 'Delta', 'Echo']
/** One paragraph per name: `Item <name> sits here.` */
const para = (name: string) => `Item ${name} sits here.`
const paragraphsDoc = () => buildDocx(NAMES.map((n) => ({ runs: [{ text: para(n) }] })))
type Policy = AuthDecision['policy']

async function session(o: Parameters<typeof start>[0] = {}) {
  const store = AuditStore.open(':memory:')
  const h = await start({ audit: new AuditTable({ store }), ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, store, stop: () => { c.close(); h.stop() } }
}

/** Sends a text op and collects its reqId's replies until one carries `last: true` or is not of `type`. */
async function collect(c: Client, msg: { reqId: number; type: string } & Record<string, unknown>, type: string) {
  c.send({ token: c.manifest!.session.token, ...msg })
  const parts: Array<{ text: Record<string, any>; bytes: number }> = []
  for (;;) {
    const m = await c.next((m) => 'text' in m && m.text.reqId === msg.reqId) as { text: Record<string, any>; bytes: number }
    if (m.text.type !== type) return { parts, end: m.text }
    parts.push(m)
    if (m.text.last === true) return { parts, end: m.text }
  }
}
const copy = async (c: Client, reqId: number, range: TextRange) => {
  const r = await collect(c, { reqId, type: 'copy', range }, 'copy-result')
  return r.end.type === 'refused' ? { refused: r.end } : { text: r.parts.map((p) => p.text.text).join(''), parts: r.parts }
}
const select = (c: Client, reqId: number, range: TextRange) => collect(c, { reqId, type: 'select', range }, 'select-result')

/** The bounding box of the first match of `query` (unique in the document; one line). */
async function rectOf(c: Client, reqId: number, query: string): Promise<{ unit: number; x: number; y: number; w: number; h: number }> {
  const { parts } = await c.text({ reqId, type: 'search', query, caseSensitive: true })
  const m = parts[0]!.text.matches[0] as { unit: number; rects: number[] }
  const q = Array.from({ length: m.rects.length / 4 }, (_, i) => m.rects.slice(i * 4, i * 4 + 4) as [number, number, number, number])
  const x = Math.min(...q.map((r) => r[0])), y = Math.min(...q.map((r) => r[1]))
  return { unit: m.unit, x, y, w: Math.max(...q.map((r) => r[0] + r[2])) - x, h: Math.max(...q.map((r) => r[1] + r[3])) - y }
}
/** A range from the start of paragraph `from` to the end of paragraph `to` (indexes into NAMES). */
async function paragraphs(c: Client, reqId: number, from: number, to: number): Promise<TextRange> {
  const a = await rectOf(c, reqId, para(NAMES[from]!)), b = await rectOf(c, reqId + 1, para(NAMES[to]!))
  return { anchor: { unit: a.unit, x: a.x + 0.5, y: a.y + a.h / 2 }, focus: { unit: b.unit, x: b.x + b.w - 0.5, y: b.y + b.h / 2 } }
}
const textOf = (from: number, to: number) => NAMES.slice(from, to + 1).map(para).join('\n')
const capped = (maxParagraphs: number, maxSessionParagraphs?: number): Policy => ({ copy: { maxParagraphs, ...(maxSessionParagraphs !== undefined ? { maxSessionParagraphs } : {}) } })

/** The local index of `bytes` (all pages), built as text-served.test.ts builds it. */
async function localIndex(bytes: Uint8Array) {
  const host = new DocumentHost()
  const o = await host.open('local', bytes, 'visible-only')
  const paintables = (host as unknown as { docs: Map<string, { paintables: Parameters<typeof buildTextIndex>[0] }> }).docs.get(o.docKey)!.paintables
  const index = await buildTextIndex([...paintables])
  host.drop(o.docKey)
  return index
}
const wholeText = (index: Awaited<ReturnType<typeof localIndex>>) => {
  const last = index.pages.length - 1, lines = index.pages[last]!.lines
  return textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: last, lineIndex: lines.length - 1, charIndex: lines[lines.length - 1]!.text.length } })
}

describe('select and copy over the wire', () => {
  test('copy off: select and copy are refused text-unavailable', async () => {
    const { c, m, stop } = await session({ doc: await paragraphsDoc() })
    expect(m.can.copy).toBe(false)
    expect((await select(c, 10, { all: true })).end).toEqual({ reqId: 10, type: 'refused', code: 'text-unavailable' })
    expect((await copy(c, 11, { all: true })).refused).toEqual({ reqId: 11, type: 'refused', code: 'text-unavailable' })
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    stop()
  }, 60_000)

  test('maxParagraphs n: a selection of n is copied, n+1 is refused copy-too-large with limit n and requested n+1', async () => {
    const { c, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2) })
    expect(await copy(c, 10, await paragraphs(c, 100, 0, 1))).toMatchObject({ text: textOf(0, 1) })
    const three = await paragraphs(c, 102, 0, 2)
    expect((await copy(c, 11, three)).refused).toEqual({ reqId: 11, type: 'refused', code: 'copy-too-large', limit: 2, requested: 3 })
    // Select is not capped: it shows the selection and its count.
    const s = await select(c, 12, three)
    expect(s.end).toMatchObject({ reqId: 12, type: 'select-result', paragraphs: 3, last: true })
    expect(s.parts.flatMap((p) => p.text.units.map((u: { unit: number }) => u.unit))).toEqual([0])
    stop()
  }, 60_000)

  test('maxSessionParagraphs: copies up to the total succeed, the next is refused copy-quota-exhausted', async () => {
    const { c, m, h, stop } = await session({ doc: await paragraphsDoc(), policy: capped(2, 5) })
    expect(m.caps.copy).toEqual({ maxParagraphs: 2, maxSessionParagraphs: 5 })
    const two = await paragraphs(c, 100, 0, 1), one = await paragraphs(c, 102, 3, 3)
    expect(await copy(c, 10, two)).toMatchObject({ text: textOf(0, 1) })
    expect(await copy(c, 11, two)).toMatchObject({ text: textOf(0, 1) })
    expect((await copy(c, 12, two)).refused).toEqual({ reqId: 12, type: 'refused', code: 'copy-quota-exhausted', limit: 5, requested: 2 })
    expect(await copy(c, 13, one)).toMatchObject({ text: para('Delta') })
    expect((await copy(c, 14, one)).refused).toEqual({ reqId: 14, type: 'refused', code: 'copy-quota-exhausted', limit: 5, requested: 1 })
    expect(h.stream.audit.get(m.session.sid)!.copied).toEqual({ events: 3, paragraphs: 5 })
    stop()
  }, 60_000)

  test('copying the same paragraph twice counts twice', async () => {
    const { c, m, h, stop } = await session({ doc: await paragraphsDoc(), policy: capped(1, 2) })
    const one = await paragraphs(c, 100, 1, 1)
    expect(await copy(c, 10, one)).toMatchObject({ text: para('Bravo') })
    expect(await copy(c, 11, one)).toMatchObject({ text: para('Bravo') })
    expect((await copy(c, 12, one)).refused).toEqual({ reqId: 12, type: 'refused', code: 'copy-quota-exhausted', limit: 2, requested: 1 })
    expect(h.stream.audit.get(m.session.sid)!.copied).toEqual({ events: 2, paragraphs: 2 })
    stop()
  }, 60_000)

  test('the session total survives a resume and is measured against the resumed cap', async () => {
    let policy: Policy = capped(3, 10)
    const { h, c, m, store, stop } = await session({ doc: await paragraphsDoc(), authorize: async () => ({ viewerId: 'alice', policy }) })
    const sid = m.session.sid
    expect(await copy(c, 10, await paragraphs(c, 100, 0, 2))).toMatchObject({ text: textOf(0, 2) })
    // Resume under a smaller session cap.
    policy = capped(3, 4)
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid, resumeProof: m.session.token } }); await b.ready()
    expect(m2.session.sid).toBe(sid)
    expect(m2.caps.copy).toEqual({ maxParagraphs: 3, maxSessionParagraphs: 4 })
    expect((await copy(b, 20, await paragraphs(b, 200, 0, 1))).refused).toEqual({ reqId: 20, type: 'refused', code: 'copy-quota-exhausted', limit: 4, requested: 2 })
    expect(await copy(b, 21, await paragraphs(b, 202, 4, 4))).toMatchObject({ text: para('Echo') })
    expect((await copy(b, 22, await paragraphs(b, 204, 4, 4))).refused).toMatchObject({ code: 'copy-quota-exhausted', limit: 4, requested: 1 })
    // The audit keeps the open-time services and every change, with when.
    const row = h.stream.audit.get(sid)!
    expect(row.services).toEqual({ copy: { maxParagraphs: 3, maxSessionParagraphs: 10 }, textA11y: 'focused', print: { maxPages: null, maxSessionPages: null } })
    expect(row.servicesChanged).toEqual([{ copy: { maxParagraphs: 3, maxSessionParagraphs: 4 }, textA11y: 'focused', print: { maxPages: null, maxSessionPages: null }, at: expect.any(Number) }])
    expect(row.copied).toEqual({ events: 2, paragraphs: 4 })
    // A resume that changes nothing writes no services event; one to 'full' does.
    const c3 = new Client(h.url)
    await c3.hello({ resume: { sid, resumeProof: m2.session.token } }); await c3.ready()
    expect(h.stream.audit.get(sid)!.servicesChanged.length).toBe(1)
    policy = { copy: 'full' }
    const c4 = new Client(h.url)
    const m4 = await c4.hello({ resume: { sid, resumeProof: c3.manifest!.session.token } }); await c4.ready()
    expect(m4.caps.copy).toEqual({ maxParagraphs: null, maxSessionParagraphs: null })
    expect(h.stream.audit.get(sid)!.servicesChanged.map((x) => [x.copy, x.textA11y])).toEqual([[{ maxParagraphs: 3, maxSessionParagraphs: 4 }, 'focused'], [{ maxParagraphs: null, maxSessionParagraphs: null }, 'unit']])
    expect(store.services(sid).length).toBe(2)
    // Under 'full' the earlier total no longer limits copying.
    expect(await copy(c4, 40, await paragraphs(c4, 400, 0, 4))).toMatchObject({ text: textOf(0, 4) })
    for (const x of [b, c3, c4]) x.close()
    stop()
  }, 60_000)

  test('a selection across two pages copies both in order and counts per page', async () => {
    // The long middle paragraph crosses the page break: it counts once per page.
    const doc = await buildDocx([{ runs: [{ text: 'Intro paragraph.' }] }, { runs: [{ text: 'word '.repeat(1500) + 'end.' }] }, { runs: [{ text: 'Closing paragraph.' }] }])
    const { c, m, store, stop } = await session({ doc, policy: { copy: 'full' } })
    expect(m.units.count).toBe(2)
    const all = await copy(c, 10, { all: true })
    const local = wholeText(await localIndex(doc))
    // Copy equals the local textForRange over both pages (no sanitise difference here).
    expect(all.text).toBe(local)
    expect(all.text!.startsWith('Intro paragraph.\nword word')).toBe(true)
    expect(all.text!.endsWith('word end.\nClosing paragraph.')).toBe(true)
    const page = (u: number): TextRange => ({ anchor: { unit: u, x: 0, y: 0 }, focus: { unit: u, x: m.units.list[u]!.w, y: m.units.list[u]!.h } })
    const counts = []
    for (const u of [0, 1]) counts.push((await select(c, 20 + u, page(u))).end.paragraphs)
    expect(counts).toEqual([2, 2])
    // Across the break: from the intro on page 0 to the closing line on page 1.
    const a = await rectOf(c, 30, 'Intro paragraph.'), b = await rectOf(c, 31, 'Closing paragraph.')
    expect([a.unit, b.unit]).toEqual([0, 1])
    const across: TextRange = { anchor: { unit: 0, x: a.x + 0.5, y: a.y + a.h / 2 }, focus: { unit: 1, x: b.x + b.w - 0.5, y: b.y + b.h / 2 } }
    const s = await select(c, 32, across)
    // Many rects (one per word): the answer spans several chunks, units in order.
    const seen = s.parts.flatMap((p) => p.text.units.map((u: { unit: number }) => u.unit))
    expect([...new Set(seen)]).toEqual([0, 1])
    expect(seen).toEqual([...seen].sort())
    for (const p of s.parts) expect(p.bytes).toBeLessThanOrEqual(KIB48)
    for (const p of s.parts) expect(p.text.paragraphs).toBe(4)
    expect(s.parts.map((p) => p.text.last)).toEqual([...s.parts.slice(1).map(() => undefined), true])
    expect(await copy(c, 33, across)).toMatchObject({ text: local })
    // The copy is charged per page too, and its record spans both units.
    expect(store.copies(m.session.sid).at(-1)).toEqual({ sid: m.session.sid, kind: 'copy', paragraphs: 4, fromUnit: 0, toUnit: 1, at: expect.any(Number) })
    stop()
  }, 60_000)

  test('rtl/mixed text copies in logical order, equal to local textForRange', async () => {
    // A generated mixed-direction document: every span is served.
    const lines = ['Mixed: שלום עולם and مرحبا بالعالم 123 end.', 'עברית בלבד כאן']
    const mixed = await buildDocx(lines.map((text) => ({ runs: [{ text }] })))
    {
      const { c, stop } = await session({ doc: mixed, policy: { copy: 'full' } })
      const r = await copy(c, 10, { all: true })
      expect(r.text).toBe(lines.join('\n'))
      expect(r.text).toBe(wholeText(await localIndex(mixed)))
      stop()
    }
    // poi-rtl.docx withholds 12 spans on its first lines (fail closed), so it is
    // compared over its fully served lines 2..5, through the same points.
    const rtl = corpus('poi-rtl.docx'), local = await localIndex(rtl)
    const { c, m, stop } = await session({ doc: rtl, policy: { copy: 'full' } })
    const lineY = (i: number) => { const s = local.pages[0]!.lines[i]!.spans[0]!; return s.y - s.fontSize / 3 }
    const anchor: UnitPoint = { unit: 0, x: 1, y: lineY(2) }, focus: UnitPoint = { unit: 0, x: m.units.list[0]!.w - 1, y: lineY(5) }
    const hit = (p: UnitPoint) => hitTest(local, 0, p.x, p.y, { strict: false })!
    const expected = textForRange(local, normalizeRange(hit(anchor), hit(focus)))
    for (const i of [3, 4]) expect(expected).toContain(local.pages[0]!.lines[i]!.text)
    expect(expected).not.toContain(local.pages[0]!.lines[0]!.text)
    expect(await copy(c, 20, { anchor, focus })).toMatchObject({ text: expected })
    stop()
  }, 60_000)

  test('a refused copy sends no copy-result and writes no audit row', async () => {
    const { c, m, store, stop } = await session({ doc: await paragraphsDoc(), policy: capped(1) })
    const two = await paragraphs(c, 100, 0, 1)
    expect((await copy(c, 10, two)).refused).toMatchObject({ code: 'copy-too-large', limit: 1, requested: 2 })
    await Bun.sleep(200)
    expect(c.msgs.filter((x) => 'text' in x && x.text.reqId === 10)).toEqual([])
    expect(store.copies(m.session.sid)).toEqual([])
    // An empty selection copies as '' with no audit row and no quota.
    const blank: TextRange = { anchor: { unit: 0, x: 1, y: m.units.list[0]!.h - 2 }, focus: { unit: 0, x: 2, y: m.units.list[0]!.h - 1 } }
    const e = await copy(c, 11, blank)
    expect(e.parts!.map((p) => p.text)).toEqual([{ reqId: 11, type: 'copy-result', seq: 0, text: '', last: true }])
    expect(store.copies(m.session.sid)).toEqual([])
    // Malformed ranges are refused policy-refused.
    for (const [i, range] of [{ all: false }, { anchor: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 1, x: 0, y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0.5, x: 0, y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0, x: 'a', y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, null].entries()) {
      expect((await collect(c, { reqId: 50 + i, type: 'copy', range }, 'copy-result')).end, JSON.stringify(range)).toEqual({ reqId: 50 + i, type: 'refused', code: 'policy-refused' })
      expect((await select(c, 70 + i, range as never)).end).toEqual({ reqId: 70 + i, type: 'refused', code: 'policy-refused' })
    }
    expect(store.copies(m.session.sid)).toEqual([])
    stop()
  }, 60_000)

  test('each successful copy writes one audit copies row with its paragraph count', async () => {
    const { c, m, store, h, stop } = await session({ doc: await paragraphsDoc(), policy: capped(3) })
    const sid = m.session.sid
    await copy(c, 10, await paragraphs(c, 100, 0, 1))
    await copy(c, 11, await paragraphs(c, 102, 2, 4))
    await copy(c, 12, { all: true }) // 5 > 3: refused, no row
    expect(store.copies(sid)).toEqual([
      { sid, kind: 'copy', paragraphs: 2, fromUnit: 0, toUnit: 0, at: expect.any(Number) },
      { sid, kind: 'copy', paragraphs: 3, fromUnit: 0, toUnit: 0, at: expect.any(Number) },
    ])
    expect(h.stream.audit.get(sid)!.copied).toEqual({ events: 2, paragraphs: 5 })
    expect(h.stream.audit.get(sid)!.services).toEqual({ copy: { maxParagraphs: 3, maxSessionParagraphs: null }, textA11y: 'focused', print: { maxPages: null, maxSessionPages: null } })
    stop()
  }, 60_000)

  test('R5: copy is charged the text rate (refusal only); select only the message budget', async () => {
    let t = Date.now()
    const { c, stop } = await session({ doc: await paragraphsDoc(), policy: { copy: 'full' }, now: () => t })
    // 200 ms apart: well inside the message budget (20/s, 120/10 s).
    for (let i = 0; i < 59; i++) { t += 200; expect((await c.text({ reqId: 100 + i, type: 'search', query: 'Item' })).end.type).toBe('search-done') }
    t += 200; expect(await copy(c, 10, { all: true })).toMatchObject({ text: textOf(0, 4) })   // the 60th text op
    for (let i = 0; i < 4; i++) { t += 200; expect((await copy(c, 20 + i, { all: true })).refused).toEqual({ reqId: 20 + i, type: 'refused', code: 'rate-limited' }) }
    t += 200; expect((await select(c, 30, { all: true })).end).toMatchObject({ reqId: 30, type: 'select-result', paragraphs: 5, last: true })
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    stop()
  }, 120_000)

  test('{all:true} on a full-copy session returns the whole document text chunked ≤ 48 KiB', async () => {
    const line = 'alpha beta gamma delta epsilon zeta eta theta "quoted"'
    const doc = await buildDocx(Array.from({ length: 1200 }, () => ({ runs: [{ text: line }] })))
    const { c, m, stop } = await session({ doc, policy: { copy: 'full' } })
    const r = await copy(c, 10, { all: true })
    expect(r.parts!.length).toBeGreaterThan(1)
    for (const p of r.parts!) expect(p.bytes).toBeLessThanOrEqual(KIB48)
    expect(r.parts!.map((p) => p.text.seq)).toEqual(r.parts!.map((_, i) => i))
    expect(r.parts!.map((p) => p.text.last)).toEqual([...r.parts!.slice(1).map(() => false), true])
    expect(r.text!.length).toBeGreaterThan(KIB48)
    // Reassembled, the chunks are the worker's whole-document answer.
    const ref = await pool().text('ref', { docRef: m.docRef, contentPolicy: 'visible-only', op: 'copy', range: { all: true }, maxParagraphs: null, maxBytes: COPY_MAX_BYTES }) as { text: string }
    expect(r.text).toBe(ref.text)
    expect(r.text!.split('\n').filter((l) => l === line).length).toBe(1200)
    stop()
  }, 120_000)
})

describe('ops.ts with a fake pool', () => {
  const slot = (): TextSlot => ({ copiedParagraphs: 0 })
  const all: TextRange = { all: true }
  function copyCtx(o: Partial<CopyContext> & Pick<CopyContext, 'job'>) {
    const sent: Array<Record<string, any>> = [], refused: Array<[string, object | undefined]> = [], records: object[] = []
    const c: CopyContext = { slot: slot(), reqId: 7, range: all, unitCount: 3, caps: () => ({ maxParagraphs: null, maxSessionParagraphs: null }), record: (r) => { records.push(r) }, alive: () => true, send: (m) => sent.push(m as never), refuse: (code, nums) => refused.push([code, nums]), ...o }
    return { c, sent, refused, records }
  }

  test('the worker job carries maxParagraphs and the 4 MiB byte cap; the record carries the unit span', async () => {
    const jobs: TextJob[] = []
    const x = copyCtx({ range: { anchor: { unit: 2, x: 1, y: 1 }, focus: { unit: 1, x: 0, y: 0 } }, caps: () => ({ maxParagraphs: 4, maxSessionParagraphs: null }), job: async (j) => { jobs.push(j); return { op: 'copy', text: 'hi', paragraphs: 1 } } })
    await runCopy(x.c)
    expect(jobs).toEqual([{ op: 'copy', range: { anchor: { unit: 2, x: 1, y: 1 }, focus: { unit: 1, x: 0, y: 0 } }, maxParagraphs: 4, maxBytes: 4194304 }])
    expect(x.records).toEqual([{ paragraphs: 1, fromUnit: 1, toUnit: 2 }])
    expect(x.sent).toEqual([{ reqId: 7, type: 'copy-result', seq: 0, text: 'hi', last: true }])
    expect(x.c.slot.copiedParagraphs).toBe(1)
  })

  test('over 4 MiB is refused with the byte cap; TextTooLarge states the cap without requested', async () => {
    const big = copyCtx({ job: async () => ({ op: 'copy', refused: 'too-large', paragraphs: 9, bytes: COPY_MAX_BYTES + 1 }) })
    await runCopy(big.c)
    expect(big.refused).toEqual([['copy-too-large', { limit: 4194304, requested: 4194305 }]])
    const huge = copyCtx({ job: async () => { throw new TextTooLarge([1, 2]) } })
    await runCopy(huge.c)
    expect(huge.refused).toEqual([['copy-too-large', { limit: 4194304 }]])
    // The main thread re-checks the worker's answer too (fail closed).
    const lying = copyCtx({ caps: () => ({ maxParagraphs: 2, maxSessionParagraphs: null }), job: async () => ({ op: 'copy', text: 'x'.repeat(COPY_MAX_BYTES + 1), paragraphs: 3 }) })
    await runCopy(lying.c)
    expect(lying.refused).toEqual([['copy-too-large', { limit: 2, requested: 3 }]])
    for (const x of [big, huge, lying]) expect([x.sent, x.records, x.c.slot.copiedParagraphs]).toEqual([[], [], 0])
  })

  test('pool errors: AdmissionRefused is rate-limited, others policy-refused', async () => {
    const busy = copyCtx({ job: async () => { throw new AdmissionRefused('session-queue') } })
    await runCopy(busy.c)
    const broken = copyCtx({ job: async () => { throw new Error('worker died') } })
    await runCopy(broken.c)
    const odd = copyCtx({ job: async () => ({ missing: [1] }) })
    await runCopy(odd.c)
    expect([busy.refused, broken.refused, odd.refused]).toEqual([[['rate-limited', undefined]], [['policy-refused', undefined]], [['policy-refused', undefined]]])
  })

  test('copy turned off by the time the copy runs is text-unavailable', async () => {
    const x = copyCtx({ caps: () => false, job: async () => { throw new Error('unreachable') } })
    await runCopy(x.c)
    expect(x.refused).toEqual([['text-unavailable', undefined]])
  })

  test('an audit write failure refuses policy-refused and sends nothing', async () => {
    const x = copyCtx({ record: () => { throw new Error('disk full') }, job: async () => ({ op: 'copy', text: 'secret', paragraphs: 1 }) })
    await runCopy(x.c)
    expect([x.refused, x.sent, x.c.slot.copiedParagraphs]).toEqual([[['policy-refused', undefined]], [], 0])
  })

  test('an empty selection: one empty copy-result, no record, no quota', async () => {
    const x = copyCtx({ caps: () => ({ maxParagraphs: 1, maxSessionParagraphs: 1 }), job: async () => ({ op: 'copy', text: '', paragraphs: 0 }) })
    x.c.slot.copiedParagraphs = 1
    await runCopy(x.c)
    expect([x.sent, x.records, x.refused, x.c.slot.copiedParagraphs]).toEqual([[{ reqId: 7, type: 'copy-result', seq: 0, text: '', last: true }], [], [], 1])
  })

  test('copies on one sid are serialized and each sees the caps in force when it runs', async () => {
    const s = slot(), order: string[] = []
    let release!: () => void
    const gate = new Promise<void>((r) => { release = r })
    let cap = 3
    const sent: Array<Record<string, any>> = [], refused: unknown[] = []
    const mk = (reqId: number, wait: boolean): CopyContext => ({
      slot: s, reqId, range: all, unitCount: 1, caps: () => ({ maxParagraphs: null, maxSessionParagraphs: cap }), record: () => {}, alive: () => true,
      send: (m) => sent.push(m as never), refuse: (code, nums) => refused.push([reqId, code, nums]),
      job: async () => { order.push(`start${reqId}`); if (wait) await gate; order.push(`end${reqId}`); return { op: 'copy', text: `t${reqId}`, paragraphs: 2 } },
    })
    const a = runCopy(mk(1, true)), b = runCopy(mk(2, false))
    await Bun.sleep(5)
    expect(order).toEqual(['start1'])
    cap = 4
    release(); await Promise.all([a, b])
    expect(order).toEqual(['start1', 'end1', 'start2', 'end2'])
    expect(sent.map((m) => m.reqId)).toEqual([1, 2])
    expect(s.copiedParagraphs).toBe(4)
    // A third copy is over the (current) cap.
    await runCopy(mk(3, false))
    expect(refused).toEqual([[3, 'copy-quota-exhausted', { limit: 4, requested: 2 }]])
  })

  test('a dead owner gets nothing and a failed copy does not break the queue', async () => {
    const s = slot()
    const dead = copyCtx({ alive: () => false, job: async () => ({ op: 'copy', text: 'x', paragraphs: 1 }) }); dead.c.slot = s
    const boom = copyCtx({ job: async () => ({ op: 'copy', text: 'x', paragraphs: 1 }), send: () => { throw new Error('socket gone') } }); boom.c.slot = s
    const ok = copyCtx({ job: async () => ({ op: 'copy', text: 'y', paragraphs: 1 }) }); ok.c.slot = s
    await Promise.all([runCopy(dead.c), runCopy(boom.c), runCopy(ok.c)])
    expect([dead.sent, dead.refused, dead.records]).toEqual([[], [], []])
    expect(ok.sent).toEqual([{ reqId: 7, type: 'copy-result', seq: 0, text: 'y', last: true }])
  })

  test('a new select supersedes an in-flight one: the old reqId gets nothing', async () => {
    const s = slot(), sent: Array<Record<string, any>> = []
    let release!: () => void
    const gate = new Promise<void>((r) => { release = r })
    const reply: TextReply = { op: 'select', units: [{ unit: 0, rects: [1, 2, 3, 4] }], paragraphs: 1 }
    const a = runSelect({ slot: s, reqId: 1, range: all, alive: () => true, send: (m) => sent.push(m as never), refuse: () => sent.push({ refused: 1 }), job: async () => { await gate; return reply } })
    const b = runSelect({ slot: s, reqId: 2, range: all, alive: () => true, send: (m) => sent.push(m as never), refuse: () => sent.push({ refused: 2 }), job: async () => reply })
    await b; release(); await a
    expect(sent).toEqual([{ reqId: 2, type: 'select-result', units: [{ unit: 0, rects: [1, 2, 3, 4] }], paragraphs: 1, last: true }])
    expect(s.select).toBeUndefined()
  })

  test('select pool errors: AdmissionRefused rate-limited, TextTooLarge policy-refused', async () => {
    const out: string[] = []
    for (const e of [new AdmissionRefused('process-queue'), new TextTooLarge([0])]) await runSelect({ slot: slot(), reqId: 1, range: all, alive: () => true, send: () => out.push('sent'), refuse: (code) => out.push(code), job: async () => { throw e } })
    expect(out).toEqual(['rate-limited', 'policy-refused'])
  })

  test('rangeArg validates and rebuilds the range', () => {
    expect(rangeArg({ all: true, extra: 1 }, 2)).toEqual({ all: true })
    const pt = { unit: 1, x: 2.5, y: -3, policy: 'x' }
    expect(rangeArg({ anchor: pt, focus: { unit: 0, x: 0, y: 0 }, more: 1 }, 2)).toEqual({ anchor: { unit: 1, x: 2.5, y: -3 }, focus: { unit: 0, x: 0, y: 0 } })
    for (const bad of [undefined, null, [], { all: 1 }, { all: false }, { anchor: { unit: 2, x: 0, y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: -1, x: 0, y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0, x: Infinity, y: 0 }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0, x: 0, y: NaN }, focus: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0, x: 0, y: 0 } }, { anchor: { unit: 0, x: 0, y: 0 }, focus: 'x' }]) {
      expect(rangeArg(bad, 2), JSON.stringify(bad)).toBeUndefined()
    }
  })
})

describe('chunking', () => {
  const size = (m: unknown) => new TextEncoder().encode(JSON.stringify(m)).length

  test('chunkText: ≤ limit, seq from 0, last on the final chunk only, never splits a surrogate pair, lossless', () => {
    const text = ('a"\\\n\u0001é中😀 ').repeat(12_000) + '😀'
    const ms = chunkText(5, text)
    expect(ms.length).toBeGreaterThan(2)
    for (const m of ms) {
      expect(size(m)).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES)
      const first = m.text.charCodeAt(0), last = m.text.charCodeAt(m.text.length - 1)
      expect(first >= 0xdc00 && first <= 0xdfff).toBe(false)
      expect(last >= 0xd800 && last <= 0xdbff).toBe(false)
    }
    expect(ms.map((m) => m.seq)).toEqual(ms.map((_, i) => i))
    expect(ms.map((m) => m.last)).toEqual([...ms.slice(1).map(() => false), true])
    expect(ms.map((m) => m.text).join('')).toBe(text)
    // Tight limits: every chunk is exactly at or under it.
    for (const limit of [80, 97, 200]) for (const m of chunkText(123456, text.slice(0, 500), limit)) expect(size(m)).toBeLessThanOrEqual(limit)
    expect(chunkText(5, '')).toEqual([{ reqId: 5, type: 'copy-result', seq: 0, text: '', last: true }])
  })

  test('chunkSelect: whole units where possible; one unit alone over the limit is split by rect; every chunk has paragraphs', () => {
    const rects = (n: number, k: number) => Array.from({ length: 4 * n }, (_, i) => k + i * 1.0000001)
    const units = [{ unit: 0, rects: rects(10, 0) }, { unit: 1, rects: rects(4000, 1) }, { unit: 2, rects: rects(3, 2) }]
    const ms = chunkSelect(9, units, 4)
    expect(ms.length).toBeGreaterThan(2)
    for (const m of ms) { expect(size(m)).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES); expect(m.paragraphs).toBe(4); for (const u of m.units) expect(u.rects.length % 4).toBe(0) }
    expect(ms.map((m) => m.last)).toEqual([...ms.slice(1).map(() => undefined), true])
    // Reassembled per unit, in order.
    const back = new Map<number, number[]>()
    for (const m of ms) for (const u of m.units) back.set(u.unit, [...(back.get(u.unit) ?? []), ...u.rects])
    expect([...back].map(([unit, r]) => ({ unit, rects: r }))).toEqual(units)
    // Small units share one message.
    expect(chunkSelect(9, [units[0]!, units[2]!], 1)).toEqual([{ reqId: 9, type: 'select-result', units: [units[0]!, units[2]!], paragraphs: 1, last: true }])
    expect(chunkSelect(9, [], 0)).toEqual([{ reqId: 9, type: 'select-result', units: [], paragraphs: 0, last: true }])
  })
})
