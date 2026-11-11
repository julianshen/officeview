/** S3a Task 1: copy / textA11y policy, manifest caps, and the resume rule (R3, R7). */
import { afterAll, describe, expect, test } from 'bun:test'
import { Client, disposePool, start } from './harness'
import { effectivePolicy, type AuthDecision } from '../src/hooks'

afterAll(() => disposePool())

async function open(policy: AuthDecision['policy']) {
  const h = await start({ policy })
  const c = new Client(h.url), m = await c.hello()
  c.close(); h.stop()
  return m
}

describe('copy / textA11y policy in the manifest', () => {
  test('defaults: search on, copy off, a11y off', async () => {
    const m = await open({})
    expect(m.can).toMatchObject({ search: true, copy: false, textA11y: false })
    expect(m.caps).toEqual({ copy: null, a11y: 'off' })
  })
  test("copy 'full' ⇒ a11y unit", async () => {
    const m = await open({ copy: 'full' })
    expect(m.can).toMatchObject({ search: true, copy: true, textA11y: true })
    expect(m.caps).toEqual({ copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit' })
  })
  test('copy cap ⇒ a11y focused', async () => {
    const m = await open({ copy: { maxParagraphs: 3 } })
    expect(m.can).toMatchObject({ copy: true, textA11y: true })
    expect(m.caps).toEqual({ copy: { maxParagraphs: 3, maxSessionParagraphs: null }, a11y: 'focused' })
  })
  test('copy cap with a session cap carries both', async () => {
    const m = await open({ copy: { maxParagraphs: 3, maxSessionParagraphs: 20 } })
    expect(m.caps).toEqual({ copy: { maxParagraphs: 3, maxSessionParagraphs: 20 }, a11y: 'focused' })
  })
  test('textA11y:true with copy off ⇒ unit', async () => {
    const m = await open({ textA11y: true })
    expect(m.can).toMatchObject({ copy: false, textA11y: true })
    expect(m.caps).toEqual({ copy: null, a11y: 'unit' })
  })
  test('textA11y:false with copy full ⇒ off', async () => {
    const m = await open({ copy: 'full', textA11y: false })
    expect(m.can).toMatchObject({ copy: true, textA11y: false })
    expect(m.caps).toEqual({ copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'off' })
  })
  test('copy or a11y on does not mark the session relaxed (R4)', async () => {
    const m = await open({ copy: 'full', textA11y: true })
    expect(m.notices).not.toContain('relaxed-session')
    expect(effectivePolicy({ policy: { copy: { maxParagraphs: 2 } } }).relaxed).toBe(false)
  })
})

describe('bad caps fail closed', () => {
  test.each([0, -1, 1.5, NaN, '3', Number.MAX_SAFE_INTEGER + 1, Infinity])('bad cap %p ⇒ copy off', (cap) => {
    const p = effectivePolicy({ policy: { copy: { maxParagraphs: cap as never } } })
    expect(p.copy).toBe(false)
    expect(p.textA11y).toBe('off')
  })
  test.each([0, -1, 1.5, NaN, '3'])('bad session cap %p ⇒ copy off', (cap) => {
    expect(effectivePolicy({ policy: { copy: { maxParagraphs: 3, maxSessionParagraphs: cap as never } } }).copy).toBe(false)
  })
  test.each([true, 'partial', null, 1, {}])('unknown copy value %p ⇒ copy off', (v) => {
    expect(effectivePolicy({ policy: { copy: v as never } }).copy).toBe(false)
  })
  test('a bad cap reaches the manifest as copy off', async () => {
    const m = await open({ copy: { maxParagraphs: -1 } })
    expect(m.can.copy).toBe(false)
    expect(m.caps).toEqual({ copy: null, a11y: 'off' })
  })
})

describe('resume (R3): copy and textA11y are mutable, everything else exact-match', () => {
  test('resume with a changed copy cap succeeds and the new manifest carries it', async () => {
    let cap = 3
    const h = await start({ authorize: async () => ({ viewerId: 'alice', policy: { copy: { maxParagraphs: cap } } }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    expect(m.caps.copy).toEqual({ maxParagraphs: 3, maxSessionParagraphs: null })
    cap = 1
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect(m2.session.sid).toBe(m.session.sid)
    expect(m2.caps).toEqual({ copy: { maxParagraphs: 1, maxSessionParagraphs: null }, a11y: 'focused' })
    expect(h.stream.table.get(m.session.sid)!.effectivePolicy.copy).toEqual({ maxParagraphs: 1, maxSessionParagraphs: null })
    expect((await a.closed).code).toBe(4409)
    b.close(); h.stop()
  })
  test('resume may turn copy off and a11y on', async () => {
    let policy: AuthDecision['policy'] = { copy: 'full' }
    const h = await start({ authorize: async () => ({ viewerId: 'alice', policy }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    policy = { textA11y: true }
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect(m2.can).toMatchObject({ copy: false, textA11y: true })
    expect(m2.caps).toEqual({ copy: null, a11y: 'unit' })
    b.close(); h.stop()
  })
  test('resume with a changed visibleMark still closes auth-failed', async () => {
    let policy: AuthDecision['policy'] = { copy: { maxParagraphs: 3 } }
    const h = await start({ authorize: async () => ({ viewerId: 'alice', policy }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    policy = { copy: { maxParagraphs: 1 }, visibleMark: 'viewer' }
    const b = new Client(h.url); await b.opened()
    b.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect((await b.closed).code).toBe(4003)
    expect(a.ws.readyState).toBe(WebSocket.OPEN)
    expect(h.stream.table.get(m.session.sid)!.effectivePolicy.copy).toEqual({ maxParagraphs: 3, maxSessionParagraphs: null })
    a.close(); h.stop()
  })
})
