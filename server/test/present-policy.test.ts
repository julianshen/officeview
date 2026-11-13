/** Stage 3c Task 1 — `can.present`, `can.notes` (R3: fail closed) and hidden slide flags. */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { canNotes, effectivePolicy, type AuthDecision } from '../src/hooks'
import { hiddenSlideDeck } from './fixtures'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

const presenter = (policy: AuthDecision['policy'] = {}): AuthDecision => ({ viewerId: 'alice', role: 'presenter', policy })

describe('policy.notes (R3: only undefined or true is on)', () => {
  test.each([[undefined], [true]])('%p is on for a presenter', (notes) => {
    const p = effectivePolicy(presenter({ notes }))
    expect(p.notes).toBe(true)
    expect(canNotes(p)).toBe(true)
  })
  test.each([[false], [null], ['yes'], [1], [0], [{}], [[]]])('%p is off', (notes) => {
    const p = effectivePolicy(presenter({ notes: notes as never }))
    expect(p.notes).toBe(false)
    expect(canNotes(p)).toBe(false)
  })
  test('an audience never gets notes', () => {
    expect(canNotes(effectivePolicy({ viewerId: 'bob', policy: { notes: true } }))).toBe(false)
    expect(canNotes(effectivePolicy({ viewerId: 'bob', role: 'audience', policy: {} }))).toBe(false)
  })
})

describe('the manifest carries can.present and can.notes', () => {
  test('a pptx session can present; a docx session cannot', async () => {
    const h = await start(), c = new Client(h.url)
    expect((await c.hello()).can.present).toBe(true)
    c.close(); h.stop()
    const d = await start({ doc: await buildDocx([{ runs: [{ text: 'Hello' }] }]) }), e = new Client(d.url)
    const m = await e.hello()
    expect(m.kind).toBe('docx')
    expect(m.can.present).toBe(false)
    e.close(); d.stop()
  })
  test.each([
    ['audience', { viewerId: 'alice', policy: {} } as AuthDecision, false],
    ['presenter', presenter(), true],
    ['presenter, notes: true', presenter({ notes: true }), true],
    ['presenter, notes: false', presenter({ notes: false }), false],
    ['presenter, notes: "yes"', presenter({ notes: 'yes' as never }), false],
    ['presenter, notes: 1', presenter({ notes: 1 as never }), false],
    ['presenter, notes: null', presenter({ notes: null as never }), false],
  ])('%s → can.notes %p', async (_, decision, notes) => {
    const h = await start({ authorize: async () => decision }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.can.notes).toBe(notes)
    expect(h.stream.audit.get(m.session.sid)!.services.notes).toBe(notes)
    c.close(); h.stop()
  })
})

describe('resume: notes is mutable, role is not', () => {
  test('a presenter resumed with notes: false gets can.notes false and the change is audited', async () => {
    let decision = presenter()
    const h = await start({ authorize: async () => decision })
    const a = new Client(h.url), m = await a.hello(); await a.ready()
    expect(m.can.notes).toBe(true)
    decision = presenter({ notes: false })
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect(m2.session.sid).toBe(m.session.sid)
    expect(m2.can.notes).toBe(false)
    expect(m2.can.present).toBe(true)
    const row = h.stream.audit.get(m.session.sid)!
    expect(row.services.notes).toBe(true)
    expect(row.servicesChanged).toEqual([{ copy: false, textA11y: 'off', print: { maxPages: null, maxSessionPages: null }, notes: false, at: expect.any(Number) }])
    // Turning it back on is a change too; a resume that changes nothing records nothing.
    decision = presenter()
    const c = new Client(h.url)
    expect((await c.hello({ resume: { sid: m.session.sid, resumeProof: m2.session.token } })).can.notes).toBe(true)
    const d = new Client(h.url)
    await d.hello({ resume: { sid: m.session.sid, resumeProof: c.manifest!.session.token } })
    expect(h.stream.audit.get(m.session.sid)!.servicesChanged.map((x) => x.notes)).toEqual([false, true])
    for (const x of [a, b, c, d]) x.close()
    h.stop()
  })
  test('a role change on resume still closes auth-failed', async () => {
    let decision = presenter()
    const h = await start({ authorize: async () => decision })
    const a = new Client(h.url), m = await a.hello(); await a.ready()
    decision = { viewerId: 'alice', role: 'audience', policy: {} }
    const b = new Client(h.url); await b.opened()
    b.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect((await b.closed).code).toBe(4003)
    expect(a.ws.readyState).toBe(WebSocket.OPEN)
    expect(h.stream.audit.get(m.session.sid)!.servicesChanged).toEqual([])
    a.close(); h.stop()
  })
})

describe('hidden slides', () => {
  test("'full': the hidden slide is a unit flagged hidden; the others carry no hidden key", async () => {
    const h = await start({ doc: await hiddenSlideDeck(), policy: { contentPolicy: 'full' } }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.units.count).toBe(3)
    expect(m.units.list[1]).toMatchObject({ hidden: true })
    expect('hidden' in m.units.list[0]!).toBe(false)
    expect('hidden' in m.units.list[2]!).toBe(false)
    expect([...h.stream.hiddenOf(m.session.sid)!]).toEqual([1])
    c.close(); h.stop()
  })
  test("'visible-only': the hidden slide is not a unit and no unit carries a hidden key", async () => {
    const h = await start({ doc: await hiddenSlideDeck() }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.units.count).toBe(2)
    expect(m.units.list.some((u) => 'hidden' in u)).toBe(false)
    expect(h.stream.hiddenOf(m.session.sid)!.size).toBe(0)
    c.close(); h.stop()
  })
  test('a deck without hidden slides flags nothing', async () => {
    const h = await start({ policy: { contentPolicy: 'full' } }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.units.list.some((u) => 'hidden' in u)).toBe(false)
    c.close(); h.stop()
  })
})
