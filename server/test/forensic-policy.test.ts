/** Stage 2a Task 4: the forensic mark is on by default; notices and the audit
 * row record what is applied (§6.1, §13.2, §13.4). */
import { afterAll, describe, expect, it } from 'bun:test'
import { AuditTable } from '../src/audit'
import { effectivePolicy, forensicNotices } from '../src/hooks'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())

describe('forensic policy', () => {
  it('effective policy: forensic on by default, off only when asked', () => {
    expect(effectivePolicy({ policy: {} })).toMatchObject({ forensicMark: true, forensicRequested: true, relaxed: false })
    expect(effectivePolicy({ policy: { forensicMark: false } })).toMatchObject({ forensicMark: false, relaxed: true })
    expect(forensicNotices(effectivePolicy({ policy: {} }))).toEqual([])
    expect(forensicNotices(effectivePolicy({ policy: { projection: 'relaxed' } }))).toEqual(['forensic-commands-unmarked'])
    expect(forensicNotices(effectivePolicy({ policy: { forensicMark: false } }))).toEqual(['forensic-off'])
  })
  it('default HELLO: no forensic-* notice; token 0 issued; l1 recorded only once a marked unit is delivered', async () => {
    const h = await start()
    const c = new Client(h.url), m = await c.hello()
    expect(m.notices.filter((n) => n.startsWith('forensic'))).toEqual([])
    const row = h.stream.audit.get(m.session.sid)!
    expect(row).toMatchObject({ forensicLayers: [], markSettings: { forensic: false, forensicRequested: true }, forensic: { codec: 'l1-bch512x40@1', token24: 0 } })
    expect(row.forensic!.msg40).toMatch(/^[0-9a-f]{10}$/)
    await c.ready(); c.unit(10, 0); await c.unitMessages(10)
    expect(h.stream.audit.get(m.session.sid)).toMatchObject({ forensicLayers: ['l1'], markSettings: { forensic: true }, untraceable: false, delivered: { marked: 1, unmarked: 0 } })
    c.close(); h.stop()
  })
  it('relaxed + visible off: command units only → the row stays untraceable despite an issued code', async () => {
    const h = await start({ doc: corpus('pypptx-test.pptx'), policy: { projection: 'relaxed', visibleMark: 'off' } })
    const c = new Client(h.url), m = await c.hello()
    await c.ready(); c.unit(10, 0)
    const { unit } = await c.unitMessages(10)
    expect(unit.discriminator).toBe(1)
    expect(h.stream.audit.get(m.session.sid)).toMatchObject({ forensic: { token24: 0 }, forensicLayers: [], untraceable: true })
    c.close(); h.stop()
  })
  it('relaxed projection: forensic-commands-unmarked', async () => {
    const h = await start({ doc: corpus('pypptx-test.pptx'), policy: { projection: 'relaxed' } })
    const c = new Client(h.url), m = await c.hello()
    expect(m.notices).toContain('forensic-commands-unmarked')
    expect(m.notices).toContain('relaxed-session')
    c.close(); h.stop()
  })
  it('forensic off + visible off: forensic-off notice and an untraceable row with no token', async () => {
    const h = await start({ policy: { forensicMark: false, visibleMark: 'off' } })
    const c = new Client(h.url), m = await c.hello()
    expect(m.notices).toContain('forensic-off')
    expect(h.stream.audit.get(m.session.sid)).toMatchObject({ forensic: null, forensicLayers: [], markSettings: { forensic: false, forensicRequested: false }, untraceable: true })
    c.close(); h.stop()
  })
  it('token exhaustion refuses a marked session with 4009; a forensic-off session still opens', async () => {
    const h = await start({ audit: new AuditTable({ firstToken: 2 ** 24 }) })
    const c = new Client(h.url)
    await c.opened(); c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1] })
    expect((await c.closed).code).toBe(4009)
    h.stop()
    const h2 = await start({ audit: new AuditTable({ firstToken: 2 ** 24 }), policy: { forensicMark: false } })
    const c2 = new Client(h2.url), m2 = await c2.hello()
    expect(m2.notices).toContain('forensic-off')
    c2.close(); h2.stop()
  })
  it('a resume keeps the original token and issues no new row', async () => {
    const h = await start({ doc: corpus('pypptx-test.pptx') })
    const c0 = new Client(h.url), m0 = await c0.hello(); await c0.ready()
    const before = h.stream.forensicOf(m0.session.sid)!, next = h.stream.audit.nextToken(1)
    const ws = c0.ws as WebSocket & { terminate?: () => void }
    if (ws.terminate) ws.terminate(); else ws.close(1001)
    await c0.closed; await Bun.sleep(50)
    const c = new Client(h.url), m = await c.hello({ resume: { sid: m0.session.sid, resumeProof: m0.session.token } })
    expect(m.session.sid).toBe(m0.session.sid)
    expect(h.stream.forensicOf(m.session.sid)).toBe(before)
    expect(h.stream.audit.nextToken(1)).toBe(next)
    c.close(); h.stop()
  })
})
