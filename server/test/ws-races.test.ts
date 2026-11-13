/** Codex review checkpoint 2 regressions: async races, resume re-evaluation,
 * session-owned rates, per-kind caps with raster fallback. */
import { afterAll, describe, expect, it } from 'bun:test'
import { AssetTable } from '../../src/stream/record/recorder'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import type { RenderedUnit } from '../src/render'
import type { Pool } from '../src/pool'
import { Client, disposePool, pool, start } from './harness'

afterAll(() => disposePool())
const settle = (ms: number) => Bun.sleep(ms)

describe('async open and close (review r2-1)', () => {
  it('a socket closed while the document is parsing never becomes a session; its document is released', async () => {
    const real = pool()
    let forgot = 0
    const slow = Object.assign(Object.create(real), {
      open: async (...a: Parameters<Pool['open']>) => { await settle(300); return real.open(...a) },
      forget: (d: string) => { forgot++; real.forget(d) },
    }) as Pool
    const h = await start({ pool: slow }), c = new Client(h.url); await c.opened()
    c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1] })
    await settle(100); c.close(); await c.closed
    await settle(600)
    expect(h.stream.stats()).toMatchObject({ sockets: 0, sessions: 0, docs: 0 })
    expect(forgot).toBe(1)
    h.stop()
  })
  it('a socket closed while authorize is pending never opens anything', async () => {
    let resolved = 0
    const h = await start({ authorize: async () => { await settle(300); return { viewerId: 'a', policy: {} } }, resolveDocument: async () => { resolved++; return null } })
    const c = new Client(h.url); await c.opened()
    c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1] })
    await settle(50); c.close(); await settle(500)
    expect(resolved).toBe(0)
    expect(h.stream.stats()).toMatchObject({ sockets: 0, sessions: 0, docs: 0 })
    h.stop()
  })
})

describe('resume re-evaluates (review r2-2/r2-6)', () => {
  it('resume always calls authorize; a revoked viewer or changed policy is refused', async () => {
    let calls = 0, mode: 'ok' | 'revoked' | 'viewer-mark' = 'ok'
    const h = await start({ authorize: async () => { calls++; return mode === 'revoked' ? null : { viewerId: 'alice', policy: mode === 'viewer-mark' ? { visibleMark: 'viewer' as const } : {} } } })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    for (const next of ['revoked', 'viewer-mark'] as const) {
      mode = next; const before = calls
      const b = new Client(h.url); await b.opened()
      b.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: m.session.token } })
      expect((await b.closed).code).toBe(4003)
      expect(calls).toBe(before + 1)
    }
    expect(a.ws.readyState).toBe(WebSocket.OPEN)
    a.close(); h.stop()
  })
  it('print caps may change on resume: accepted, carried in the manifest, recorded', async () => {
    let print: AuthDecision['policy']['print'] = true
    const store = AuditStore.open(':memory:')
    const h = await start({ audit: new AuditTable({ store }), authorize: async () => ({ viewerId: 'alice', policy: { print } }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    expect(m.caps.print).toEqual({ maxPages: null, maxSessionPages: null })
    print = { maxPages: 1 }
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } })
    expect(m2.session.sid).toBe(m.session.sid)
    expect(m2.can.print).toBe(true)
    expect(m2.caps.print).toEqual({ maxPages: 1, maxSessionPages: null })
    expect(store.services(m.session.sid).map((r) => r.print)).toEqual([{ maxPages: 1, maxSessionPages: null }])
    expect(h.stream.audit.get(m.session.sid)!.servicesChanged.map((r) => r.print)).toEqual([{ maxPages: 1, maxSessionPages: null }])
    // Turning it off on resume is also accepted.
    await b.ready(); print = false
    const c = new Client(h.url)
    const m3 = await c.hello({ resume: { sid: m.session.sid, resumeProof: b.manifest!.session.token } })
    expect(m3.can.print).toBe(false); expect(m3.caps.print).toBeNull()
    expect(store.services(m.session.sid).map((r) => r.print)).toEqual([{ maxPages: 1, maxSessionPages: null }, false])
    for (const x of [a, b, c]) x.close()
    h.stop()
  })
  it('a non-token proof needs the hook to grant it', async () => {
    const h = await start({ authorize: async (req) => ({ viewerId: 'alice', policy: {}, ...(req.token === 'granted-proof' ? { resumeGranted: true } : {}) }) })
    const a = new Client(h.url); const m = await a.hello(); await a.ready()
    const forged = new Client(h.url); await forged.opened()
    forged.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: 'not-a-token' } })
    expect((await forged.closed).code).toBe(4003)
    const b = new Client(h.url)
    expect((await b.hello({ resume: { sid: m.session.sid, resumeProof: 'granted-proof' } })).session.sid).toBe(m.session.sid)
    expect((await a.closed).code).toBe(4409)
    b.close(); h.stop()
  })
  it('two simultaneous resumes with one proof: exactly one wins', async () => {
    const h = await start(), a = new Client(h.url); const m = await a.hello(); await a.ready()
    const b = new Client(h.url), c = new Client(h.url); await Promise.all([b.opened(), c.opened()])
    for (const x of [b, c]) x.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1], resume: { sid: m.session.sid, resumeProof: m.session.token } })
    await settle(500)
    const manifests = [b, c].filter((x) => x.msgs.some((y) => 'binary' in y && y.binary.kind === 1))
    expect(manifests.length).toBe(1)
    const loser = manifests[0] === b ? c : b
    expect((await loser.closed).code).toBe(4003)
    expect(h.stream.stats().sessions).toBe(1)
    b.close(); c.close(); h.stop()
  })
  it('the RENEW cycle follows the resumed socket (review r2-5)', async () => {
    const h = await start({ tokenTtlSec: 2, renewLeadSec: 1 }), a = new Client(h.url); const m = await a.hello(); await a.ready()
    await settle(200)
    const renewsOn = (x: Client) => x.msgs.filter((y) => 'text' in y && y.text.type === 'renew').length
    const b = new Client(h.url); await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } }); await b.ready()
    await a.closed
    const aBefore = renewsOn(a)   // a may legitimately have renewed before it was superseded
    const renew = (await b.next((x) => 'text' in x && x.text.type === 'renew', 5000) as { text: Record<string, unknown> }).text
    expect(renewsOn(a)).toBe(aBefore)
    b.send({ reqId: 9, type: 'renew-ack', jti: renew.jti })
    await settle(300)
    expect(b.ws.readyState).toBe(WebSocket.OPEN)
    b.close(); h.stop()
  }, 15_000)
})

describe('rates belong to the session (review r2-7)', () => {
  it('a resume does not reset the per-sid unit budget', async () => {
    const h = await start(), a = new Client(h.url); const m = await a.hello(); await a.ready()
    for (let i = 0; i < 60; i++) { a.unit(1000 + i); if (i % 15 === 14) await settle(1100) }
    await a.unitMessages(1059)
    const b = new Client(h.url); const m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } }); await b.ready()
    b.unit(2000, 0, 1, m2.session.token)
    expect((await b.next((x) => 'text' in x && x.text.reqId === 2000) as { text: Record<string, unknown> }).text.code).toBe('rate-limited')
    b.close(); h.stop()
  }, 30_000)
})

describe('per-kind caps and raster fallback (review r2-8/r2-9)', () => {
  it('an asset whose ASSET_DEFINE exceeds 1 MiB encoded ships the unit as a raster', async () => {
    const real = pool()
    const noisy = new AssetTable()
    // 4 MiB of random RGBA: under ASSET_DEFINE's 8 MiB decoded cap, far over its 1 MiB encoded cap.
    const rgba = new Uint8Array(1024 * 1024 * 4)
    for (let k = 0; k < rgba.length; k += 65536) crypto.getRandomValues(rgba.subarray(k, k + 65536))
    noisy.add({ type: 'bitmap', source: 'image', w: 1024, h: 1024, rgba })
    const projections: string[] = []
    const stub = Object.assign(Object.create(real), {
      render: async (_s: string, job: { projection: string }) => {
        projections.push(job.projection)
        const raster = await real.render(_s, job as never, 1024)
        if (job.projection === 'final-raster') return raster
        return { kind: 'commands', w: 10, h: 10, ops: new Uint8Array(0), opCount: 0, assets: noisy.all(), faces: [] } satisfies RenderedUnit
      },
    }) as Pool
    const h = await start({ pool: stub, policy: { projection: 'relaxed' } }), c = new Client(h.url); await c.hello(); await c.ready()
    c.unit(5)
    const { unit, messages } = await c.unitMessages(5)
    expect(projections).toEqual(['commands', 'final-raster'])
    expect(unit.discriminator).toBe(2)
    expect(messages.map((x) => x.kind)).toEqual([3])
    c.close(); h.stop()
  }, 30_000)
  it('PR #64: assets each under 8 MiB but over the 9 MiB client allowance together ship the unit as a raster', async () => {
    const real = pool()
    const two = new AssetTable()
    // Two compressible 5.3 MiB bitmaps: each ASSET_DEFINE passes its caps, the pair exceeds ASSET_ALLOWANCE.
    for (const source of ['image', 'canvas'] as const) two.add({ type: 'bitmap', source, w: 1200, h: 1100, rgba: new Uint8Array(1200 * 1100 * 4).fill(source === 'image' ? 7 : 9) })
    const projections: string[] = []
    const stub = Object.assign(Object.create(real), {
      render: async (_s: string, job: { projection: string }) => {
        projections.push(job.projection)
        const raster = await real.render(_s, job as never, 1024)
        if (job.projection === 'final-raster') return raster
        return { kind: 'commands', w: 10, h: 10, ops: new Uint8Array(0), opCount: 0, assets: two.all(), faces: [] } satisfies RenderedUnit
      },
    }) as Pool
    const h = await start({ pool: stub, policy: { projection: 'relaxed' } }), c = new Client(h.url); await c.hello(); await c.ready()
    c.unit(6)
    const { unit, messages } = await c.unitMessages(6)
    expect(projections).toEqual(['commands', 'final-raster'])
    expect(unit.discriminator).toBe(2)
    expect(messages.map((x) => x.kind)).toEqual([3])
    c.close(); h.stop()
  }, 30_000)
})
