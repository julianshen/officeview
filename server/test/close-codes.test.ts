/** §6.2 rule 8: every registry row reached by a test (§16 T6/T10 close-code matrix). */
import { afterAll, describe, expect, it } from 'bun:test'
import JSZip from 'jszip'
import { Client, start, disposePool } from './harness'

afterAll(() => disposePool())
const odt = async () => { const z = new JSZip(); z.file('mimetype', 'application/vnd.oasis.opendocument.text'); z.file('content.xml', '<a/>'); return z.generateAsync({ type: 'uint8array' }) }
const binaries = (c: Client) => c.msgs.filter((m) => 'binary' in m).length

describe('close-code registry', () => {
  it('4000 unsupported-version: empty version intersection', async () => {
    const h = await start(), c = new Client(h.url); await c.opened()
    c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [2, 3] })
    expect(await c.closed).toEqual({ code: 4000, reason: 'unsupported-version' }); h.stop()
  })
  it('4000: a HELLO over 4 KiB, malformed JSON, and a 70 KiB message after READY', async () => {
    const h = await start()
    for (const raw of [JSON.stringify({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], pad: 'p'.repeat(5000) }), '{not json', JSON.stringify([1])]) {
      const c = new Client(h.url); await c.opened(); c.ws.send(raw)
      expect((await c.closed).code).toBe(4000)
    }
    const c = new Client(h.url); await c.hello(); await c.ready()
    c.ws.send(JSON.stringify({ reqId: 5, type: 'noop', pad: 'p'.repeat(70 * 1024) }))
    expect((await c.closed).code).toBe(4000); h.stop()
  })
  it('4001 unsupported-format: ODT refused at admission on the original bytes', async () => {
    const h = await start({ doc: await odt() }), c = new Client(h.url); await c.opened()
    c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1] })
    expect((await c.closed).code).toBe(4001); h.stop()
  })
  it('4002 ready-timeout: no READY inside the window, nothing streams', async () => {
    const h = await start({ readyMs: 300 }), c = new Client(h.url); await c.hello()
    const before = binaries(c)
    expect((await c.closed).code).toBe(4002)
    expect(binaries(c)).toBe(before); h.stop()
  })
  it('4002: a READY with the wrong jti is not a valid READY', async () => {
    const h = await start({ readyMs: 400 }), c = new Client(h.url); await c.hello()
    c.send({ reqId: 2, type: 'ready', tokenJti: 'nope' })
    expect((await c.next((m) => 'text' in m && m.text.type === 'refused') as { text: Record<string, unknown> }).text.code).toBe('policy-refused')
    expect((await c.closed).code).toBe(4002); h.stop()
  })
  it('4003 auth-failed: authorize null, document null, missing resumeProof, wrong-viewer proof', async () => {
    for (const o of [{ authorize: async () => null }, { doc: null }]) {
      const h = await start(o), c = new Client(h.url); await c.opened()
      c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1] })
      expect((await c.closed).code).toBe(4003); h.stop()
    }
    const h = await start({ authorize: async (req) => ({ viewerId: req.resume ? 'mallory' : 'alice', policy: {} }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    const noProof = new Client(h.url); await noProof.opened()
    noProof.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid } })
    expect((await noProof.closed).code).toBe(4003)
    const wrong = new Client(h.url); await wrong.opened()
    wrong.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: 'forged.by.mallory' } })
    expect((await wrong.closed).code).toBe(4003)
    expect(a.ws.readyState).toBe(WebSocket.OPEN)   // the live session is untouched
    a.close(); h.stop()
  })
  it('4004 session-expired: RENEW never acknowledged', async () => {
    const h = await start({ tokenTtlSec: 2, renewLeadSec: 1 }), c = new Client(h.url); await c.hello(); await c.ready()
    const renew = await c.next((m) => 'text' in m && m.text.type === 'renew', 5000)
    expect((renew as { text: Record<string, unknown> }).text.jti).toBeString()
    expect((await c.closed).code).toBe(4004); h.stop()
  })
  it('4005 revoked', async () => {
    const h = await start(), c = new Client(h.url); const m = await c.hello(); await c.ready()
    expect(h.stream.revoke(m.session.sid)).toBe(true)
    expect((await c.closed).code).toBe(4005); h.stop()
  })
  it('4006 layout-changed: fingerprint moves mid-session', async () => {
    let changed = false
    const real = (await import('../src/render')).layoutFingerprint
    const h = await start({ fingerprint: (id) => changed ? 'f'.repeat(64) : real(id) }), c = new Client(h.url); await c.hello(); await c.ready()
    changed = true; c.unit(10)
    expect((await c.closed).code).toBe(4006); h.stop()
  })
  it('4008 rate-limited: two refusals, then the third breach in 60 s closes', async () => {
    const h = await start(), c = new Client(h.url); await c.hello(); await c.ready()
    for (let i = 0; i < 25; i++) c.send({ reqId: 100 + i, type: 'noop' })
    expect((await c.closed).code).toBe(4008)
    const limited = c.msgs.filter((m) => 'text' in m && m.text.code === 'rate-limited')
    expect(limited.length).toBe(2); h.stop()
  })
  it('4009 policy-refused: any client relax attempt', async () => {
    const h = await start(), c = new Client(h.url); await c.hello(); await c.ready()
    c.send({ reqId: 9, type: 'relax', projection: 'commands' })
    expect((await c.closed).code).toBe(4009); h.stop()
  })
  it('4409 superseded: resume with the live token rotates the jti and closes the old socket', async () => {
    const h = await start(), a = new Client(h.url); const m = await a.hello(); await a.ready()
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect(m2.session.sid).toBe(m.session.sid)
    expect(b.jti()).not.toBe(a.jti())
    expect(await a.closed).toEqual({ code: 4409, reason: 'superseded' })
    await b.ready()
    b.unit(11)
    expect((await b.unitMessages(11)).unit.discriminator).toBe(2)
    b.unit(12, 0, 1, m.session.token)   // the superseded token
    expect((await b.next((x) => 'text' in x && x.text.reqId === 12) as { text: Record<string, unknown> }).text.code).toBe('session-rotated')
    b.close(); h.stop()
  }, 30_000)
  it('session-rotated is a refusal after an acknowledged RENEW, never a close', async () => {
    const h = await start({ tokenTtlSec: 2, renewLeadSec: 1 }), c = new Client(h.url); const m = await c.hello(); await c.ready()
    const renew = (await c.next((x) => 'text' in x && x.text.type === 'renew', 5000) as { text: Record<string, any> }).text
    c.send({ reqId: 40, type: 'renew-ack', jti: c.jti() })   // old jti: refused
    expect((await c.next((x) => 'text' in x && x.text.reqId === 40) as { text: Record<string, unknown> }).text.code).toBe('policy-refused')
    c.send({ reqId: 41, type: 'renew-ack', jti: renew.jti })
    await Bun.sleep(100)
    c.unit(42, 0, 1, m.session.token)
    expect((await c.next((x) => 'text' in x && x.text.reqId === 42) as { text: Record<string, unknown> }).text.code).toBe('session-rotated')
    c.unit(43, 0, 1, renew.token)
    expect((await c.unitMessages(43)).unit.discriminator).toBe(2)
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    c.close(); h.stop()
  }, 30_000)
})
