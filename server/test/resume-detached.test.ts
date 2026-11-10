/** 1d Task 1: a session whose socket drops abnormally (1001/1006) stays
 * resumable for a grace window, capped by its token's own exp (§6.2 rule 7). */
import { afterAll, describe, expect, it } from 'bun:test'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const settle = (ms: number) => Bun.sleep(ms)

/** Drop the socket without a close handshake from the client's side (abnormal). */
function drop(c: Client) {
  const ws = c.ws as WebSocket & { terminate?: () => void }
  if (ws.terminate) ws.terminate(); else ws.close(1001)
}
async function resume(url: string, sid: string, token: string) {
  const c = new Client(url)
  const m = await c.hello({ resume: { sid, resumeProof: token } }).catch(() => undefined)
  return { c, m }
}

describe('detached-session grace', () => {
  it('an abnormal drop keeps the session: resume inside grace has the same sid and a new jti and serves units', async () => {
    const h = await start({ doc: corpus('pypptx-test.pptx') })
    const c0 = new Client(h.url); const m0 = await c0.hello(); await c0.ready()
    const jti0 = c0.jti()
    drop(c0); await c0.closed; await settle(50)
    expect(h.stream.stats()).toMatchObject({ sockets: 0, sessions: 1, docs: 1 })
    const { c, m } = await resume(h.url, m0.session.sid, m0.session.token)
    expect(m?.session.sid).toBe(m0.session.sid)
    expect(c.jti()).not.toBe(jti0)
    await c.ready(); c.unit(5, 0)
    const u = await c.unitMessages(5)
    expect(u.unit.unitId).toBe(0)
    c.close(); h.stop()
  })
  it('after the grace window the session is gone: resume closes auth-failed and the document is released', async () => {
    const h = await start({ detachGraceMs: 200 })
    const c0 = new Client(h.url); const m0 = await c0.hello(); await c0.ready()
    drop(c0); await c0.closed; await settle(400)
    expect(h.stream.stats()).toMatchObject({ sessions: 0, docs: 0 })
    const { c } = await resume(h.url, m0.session.sid, m0.session.token)
    expect((await c.closed).code).toBe(4003)
    h.stop()
  })
  it('the grace is capped by the token exp', async () => {
    const h = await start({ tokenTtlSec: 1, renewLeadSec: 0, detachGraceMs: 30_000 })
    const c0 = new Client(h.url); const m0 = await c0.hello(); await c0.ready()
    drop(c0); await c0.closed; await settle(1500)
    expect(h.stream.stats()).toMatchObject({ sessions: 0, docs: 0 })
    const { c } = await resume(h.url, m0.session.sid, m0.session.token)
    expect((await c.closed).code).toBe(4003)
    h.stop()
  })
  it('a client close (1000) tears the session down immediately', async () => {
    const h = await start()
    const c0 = new Client(h.url); await c0.hello(); await c0.ready()
    c0.ws.close(1000); await c0.closed; await settle(50)
    expect(h.stream.stats()).toMatchObject({ sessions: 0, docs: 0 })
    h.stop()
  })
  it('a server-terminal close (revoke 4005) leaves nothing to resume', async () => {
    const h = await start()
    const c0 = new Client(h.url); const m0 = await c0.hello(); await c0.ready()
    h.stream.revoke(m0.session.sid)
    expect((await c0.closed).code).toBe(4005)
    await settle(50)
    expect(h.stream.stats()).toMatchObject({ sessions: 0, docs: 0 })
    const { c } = await resume(h.url, m0.session.sid, m0.session.token)
    expect((await c.closed).code).toBe(4003)
    h.stop()
  })
  it('revoking a detached session tears it down', async () => {
    const h = await start()
    const c0 = new Client(h.url); const m0 = await c0.hello(); await c0.ready()
    drop(c0); await c0.closed; await settle(50)
    expect(h.stream.revoke(m0.session.sid)).toBe(true)
    expect(h.stream.stats()).toMatchObject({ sessions: 0, docs: 0 })
    h.stop()
  })
})
