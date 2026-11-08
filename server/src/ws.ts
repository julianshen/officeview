/**
 * The only endpoint (§9.1): WS upgrade with an origin allowlist, the §6.2
 * state machine (HELLO → MANIFEST → READY → STREAM, RENEW, RESUME), per-op
 * token + capability checks, §14 rate gates, and the §6.2 rule-8 close-code
 * registry. No HTTP document route exists: every non-upgrade request is 404.
 */
import { gzipSync } from 'node:zlib'
import type { Server, ServerWebSocket } from 'bun'
import { encodeMessage } from '../../src/stream/frame'
import { encodeUnit } from '../../src/stream/unit'
import { encodeRasterBody } from '../../src/stream/raster'
import { AssetTooLarge, encodeAssets } from '../../src/stream/asset-wire'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { AuditTable } from './audit'
import { createRenderCache, DocumentLedger, renderKey, type RenderKeyFields } from './cache'
import { loadPack, requiredFacesOf } from './fonts'
import { effectivePolicy, FORENSIC_NOTICE, type AuthDecision, type AuthRequest, type Authorize, type EffectivePolicy, type ResolveDocument } from './hooks'
import { docRefOf, type Keys } from './keys'
import { admit } from './policy'
import { AdmissionRefused, RenderFailed, type Pool } from './pool'
import { IpRates, SessionRates } from './rate'
import { layoutFingerprint, SCALE_BUCKETS, type MarkSpec, type RenderedUnit, type UnitSize } from './render'
import { SessionTable, type Session } from './sessionTable'
import { issueToken, newSid, verifyToken } from './tokens'

/** §6.2 rule 8 — the complete registry. 4007 is client-only and never sent here. */
export const CLOSE = {
  'unsupported-version': 4000, 'unsupported-format': 4001, 'ready-timeout': 4002, 'auth-failed': 4003,
  'session-expired': 4004, 'revoked': 4005, 'layout-changed': 4006, 'rate-limited': 4008, 'policy-refused': 4009,
} as const
export type CloseName = keyof typeof CLOSE
export type RefusalCode = 'rate-limited' | 'policy-refused' | 'session-rotated'

const KIND = { MANIFEST: 1, ASSET_DEFINE: 2, UNIT_RASTER: 3, UNIT_COMMANDS: 4 } as const
const HELLO_CAP = 4 * 1024, TEXT_CAP = 64 * 1024, FRAGMENT = 1 << 20
const ENCODED_CAP = 8 * 1024 * 1024
const PROCESS_SOCKETS = 256
const enc = new TextEncoder()

export interface StreamServerOptions {
  keys: Keys
  originAllowlist: readonly string[]
  authorize: Authorize
  resolveDocument: ResolveDocument
  pool: Pool
  /** ms clock (tests inject). */
  now?: () => number
  tokenTtlSec?: number
  renewLeadSec?: number
  readyMs?: number
  helloMs?: number
  /** Current layout fingerprint for a policy id (tests inject a change). */
  fingerprint?: (contentPolicyId: string) => string
  path?: string
}

interface Conn {
  ip: string
  origin: string
  url: string
  headers: Headers
  phase: 'hello' | 'opening' | 'manifest' | 'stream' | 'closed'
  sid?: string
  rates: SessionRates
  timers: Set<ReturnType<typeof setTimeout>>
  missingFaces: string[]
}
type WS = ServerWebSocket<Conn>

interface DocState { docRef: string; name: string; bytes: Uint8Array; units: UnitSize[]; fingerprint: string; contentPolicyId: string; format: 'docx' | 'xlsx' | 'pptx'; refs: number }

const isU32 = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0 && (n as number) <= 0xffffffff

export function createStreamServer(o: StreamServerOptions) {
  const now = o.now ?? Date.now
  const nowSec = () => Math.floor(now() / 1000)
  const ttl = o.tokenTtlSec ?? 300, lead = o.renewLeadSec ?? 5
  const readyMs = o.readyMs ?? 2000, helloMs = o.helloMs ?? 5000
  const fingerprintOf = o.fingerprint ?? layoutFingerprint
  const table = new SessionTable()
  const audit = new AuditTable()
  const ipRates = new IpRates()
  const docs = new Map<string, DocState>()          // docRef → state
  const sessionDoc = new Map<string, { doc: DocState; policy: EffectivePolicy; decision: AuthDecision; tokenExp: number }>()
  const sockets = new Set<WS>()
  const renderCache = createRenderCache<RenderedUnit>()
  const ledger = new DocumentLedger((docRef, policy) => o.pool.dropParsed(docRef, policy === 'full@1' ? 'full' : 'visible-only'))
  const pack = loadPack()

  const send = (ws: WS, msg: object) => { if (ws.data.phase !== 'closed') ws.send(JSON.stringify(msg)) }
  const refuse = (ws: WS, reqId: unknown, code: RefusalCode) => send(ws, { reqId: isU32(reqId) ? reqId : 0, type: 'refused', code })
  const close = (ws: WS, name: CloseName | 'superseded') => {
    if (ws.data.phase === 'closed') return
    const code = name === 'superseded' ? 4409 : CLOSE[name]
    ws.data.phase = 'closed'
    for (const t of ws.data.timers) clearTimeout(t)
    ws.close(code, name)
  }
  const later = (ws: WS, ms: number, fn: () => void) => {
    const t = setTimeout(() => { ws.data.timers.delete(t); fn() }, Math.max(0, ms))
    ws.data.timers.add(t)
    return t
  }
  const sendBinary = (ws: WS, kind: number, reqId: number, payload: Uint8Array) => {
    const frames = encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => gzipSync(b, { level: 6 }), fragmentBytes: FRAGMENT })
    const total = frames.reduce((n, f) => n + f.length, 0)
    if (total > ENCODED_CAP + frames.length * 33) return false
    for (const f of frames) ws.sendBinary(f)
    return true
  }

  function markOf(p: EffectivePolicy, d: AuthDecision): MarkSpec {
    if (p.visibleMark === 'off') return { kind: 'off' }
    if (p.visibleMark === 'viewer') return { kind: 'viewer', text: d.viewerId ?? 'viewer' }
    return { kind: 'fixed' }
  }

  /** RENEW ≥ lead before exp; without an ack the old token dies at its own exp. */
  function scheduleRenew(ws: WS, sid: string, exp: number) {
    later(ws, (exp - lead) * 1000 - now(), async () => {
      const s = table.get(sid), sd = sessionDoc.get(sid)
      if (!s || !sd || s.activeSocket !== (ws as unknown) || ws.data.phase === 'closed') return
      const next = await issueToken(o.keys.kToken, { docRef: sd.doc.docRef, sid, ttlSec: ttl, now: nowSec() })
      table.beginRotation(sid, next.jti)
      ;(s as Session & { pendingExp?: number }).pendingExp = next.exp
      send(ws, { reqId: 0, type: 'renew', token: next.token, jti: next.jti, exp: next.exp })
      later(ws, exp * 1000 - now(), () => {
        const cur = table.get(sid)
        if (cur && cur.pendingJti === next.jti) close(ws, 'session-expired')
      })
    })
  }

  async function openSession(ws: WS, hello: Record<string, unknown>, reqId: number) {
    const req: AuthRequest = { headers: ws.data.headers, ip: ws.data.ip, url: new URL(ws.data.url), ...(typeof hello.token === 'string' ? { token: hello.token } : {}) }
    const resume = hello.resume as { sid?: unknown; resumeProof?: unknown } | undefined
    if (resume !== undefined) return resumeSession(ws, req, resume, reqId)
    const decision = await o.authorize(req)
    if (!decision) return close(ws, 'auth-failed')
    const resolved = await o.resolveDocument(req, decision)
    if (!resolved) return close(ws, 'auth-failed')
    if (!admit(resolved.bytes).ok) return close(ws, 'unsupported-format')
    const policy = effectivePolicy(decision)
    const docRef = await docRefOf(o.keys.kDoc, resolved.bytes)
    let opened
    try { opened = await o.pool.open(docRef, resolved.bytes, policy.contentPolicy) } catch (e) {
      return close(ws, (e as RenderFailed).refusal === 'unsupported-format' ? 'unsupported-format' : 'auth-failed')
    }
    if (ws.data.phase === 'closed') return
    let doc = docs.get(docRef)
    if (!doc) { doc = { docRef, name: resolved.name ?? 'document', bytes: resolved.bytes, units: opened.units, fingerprint: opened.layoutFingerprint, contentPolicyId: opened.contentPolicyId, format: opened.format, refs: 0 }; docs.set(docRef, doc) }
    doc.refs++
    ledger.touch(docRef, opened.contentPolicyId, opened.layoutFingerprint, resolved.bytes.length)
    const sid = newSid()
    const token = await issueToken(o.keys.kToken, { docRef, sid, ttlSec: ttl, now: nowSec() })
    const s = table.create({ sid, docRef, ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), role: policy.role, effectivePolicy: policy, activeSocket: ws as never, jti: token.jti, startedAt: nowSec() })
    sessionDoc.set(sid, { doc: { ...doc, units: opened.units, fingerprint: opened.layoutFingerprint, contentPolicyId: opened.contentPolicyId }, policy, decision, tokenExp: token.exp })
    audit.record(sid, { ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), ip: ws.data.ip, startedAt: s.startedAt, keyGeneration: 1, policyId: opened.contentPolicyId, markSettings: { visible: policy.visibleMark, forensic: false, forensicRequested: policy.forensicRequested }, relaxed: policy.relaxed, untraceable: policy.visibleMark === 'off' })
    ws.data.sid = sid
    sendManifest(ws, reqId, sid, token.token)
    scheduleRenew(ws, sid, token.exp)
  }

  async function resumeSession(ws: WS, req: AuthRequest, resume: { sid?: unknown; resumeProof?: unknown }, reqId: number) {
    if (typeof resume.sid !== 'string' || typeof resume.resumeProof !== 'string') return close(ws, 'auth-failed')
    const s = table.get(resume.sid), sd = sessionDoc.get(resume.sid)
    if (!s || !sd) return close(ws, 'auth-failed')
    const v = await verifyToken(o.keys.kToken, resume.resumeProof, { now: nowSec(), docRef: s.docRef, sid: s.sid })
    let ok = v.ok && ['current', 'rotating'].includes(table.acceptJti(s.sid, v.claims.jti))
    if (!ok) {
      // Or a proof the developer's authorize hook freshly issues for this viewer.
      const d = await o.authorize({ ...req, token: resume.resumeProof, resume: { sid: s.sid } })
      ok = !!d && d.viewerId === s.viewerId && JSON.stringify(effectivePolicy(d)) === JSON.stringify(s.effectivePolicy)
    }
    if (!ok || ws.data.phase === 'closed') return close(ws, 'auth-failed')
    const token = await issueToken(o.keys.kToken, { docRef: s.docRef, sid: s.sid, ttlSec: ttl, now: nowSec() })
    const old = s.activeSocket as unknown as WS
    if (old !== ws && old.data) { old.data.phase = 'closed'; for (const t of old.data.timers) clearTimeout(t) }
    table.supersede(s.sid, ws as never, token.jti)
    sd.tokenExp = token.exp
    ws.data.sid = s.sid
    sendManifest(ws, reqId, s.sid, token.token)
    scheduleRenew(ws, s.sid, token.exp)
  }

  function sendManifest(ws: WS, reqId: number, sid: string, token: string) {
    const sd = sessionDoc.get(sid)!, p = sd.policy
    const m: Manifest = {
      v: 1, docRef: sd.doc.docRef, kind: sd.doc.format, name: sd.doc.name,
      units: { count: sd.doc.units.length, list: sd.doc.units },
      layoutFingerprint: sd.doc.fingerprint,
      session: { sid, token, expiresSec: ttl, resume: 'session' },
      policy: { contentPolicy: p.contentPolicy, protection: p.projection, semantic: 'visible-only', contentPolicyId: sd.doc.contentPolicyId, viewerMarkId: p.visibleMark },
      // Stage 1: semantic services, presentation and notes land in Stage 3.
      can: { search: false, copy: false, print: p.print, textA11y: false, present: false, notes: false, links: false },
      recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [...SCALE_BUCKETS], requiredFaces: p.projection === 'commands' ? requiredFacesOf(pack.faces.map((f) => f.id), pack) : [], projection: p.projection },
      limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: FRAGMENT, tilePx: 512 },
      notices: [FORENSIC_NOTICE, ...(p.relaxed ? ['relaxed-session'] : [])],
    }
    ws.data.phase = 'manifest'
    sendBinary(ws, KIND.MANIFEST, reqId, encodeManifest(m))
    later(ws, readyMs, () => { if (ws.data.phase === 'manifest') close(ws, 'ready-timeout') })
  }

  function onReady(ws: WS, msg: Record<string, unknown>) {
    const s = ws.data.sid ? table.get(ws.data.sid) : undefined
    if (ws.data.phase !== 'manifest' || !s) return refuse(ws, msg.reqId, 'policy-refused')
    if (msg.tokenJti !== s.jti) return refuse(ws, msg.reqId, 'policy-refused')   // not a valid READY; the window keeps running
    const faces = msg.missingFaces
    if (faces !== undefined && (!Array.isArray(faces) || faces.length > 64 || faces.some((f) => typeof f !== 'string' || f.length > 64))) return refuse(ws, msg.reqId, 'policy-refused')
    ws.data.missingFaces = (faces as string[] | undefined) ?? []
    ws.data.phase = 'stream'
    s.status = 'stream'
  }

  async function onUnit(ws: WS, msg: Record<string, unknown>) {
    const reqId = msg.reqId
    const sid = ws.data.sid, s = sid ? table.get(sid) : undefined, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !s || !sd || !isU32(reqId) || reqId === 0) return refuse(ws, reqId, 'policy-refused')
    const v = await verifyToken(o.keys.kToken, msg.token, { now: nowSec(), docRef: sd.doc.docRef, sid: s.sid })
    if (!v.ok) return v.reason === 'exp' ? close(ws, 'session-expired') : refuse(ws, reqId, 'policy-refused')
    const jti = table.acceptJti(s.sid, v.claims.jti)
    if (jti === 'rotated') return refuse(ws, reqId, 'session-rotated')
    if (jti === 'unknown') return refuse(ws, reqId, 'policy-refused')
    const unitIndex = msg.unitIndex, scale = msg.scale
    if (!Number.isInteger(unitIndex) || (unitIndex as number) < 0 || (unitIndex as number) >= sd.doc.units.length) return refuse(ws, reqId, 'policy-refused')
    if (!(SCALE_BUCKETS as readonly unknown[]).includes(scale)) return refuse(ws, reqId, 'policy-refused')
    const rate = ws.data.rates.check('unit', now())
    if (rate === 'close') return close(ws, 'rate-limited')
    if (rate === 'refuse') return refuse(ws, reqId, 'rate-limited')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    const p = sd.policy, mark = markOf(p, sd.decision)
    const missing = p.projection === 'commands' ? [...ws.data.missingFaces].sort() : undefined
    const fields: RenderKeyFields = {
      docRef: sd.doc.docRef, contentPolicyId: sd.doc.contentPolicyId, layoutFingerprint: sd.doc.fingerprint, recorderVersion: 1, scaleBucket: scale as number,
      markSettings: { visible: mark.kind, text: mark.kind === 'viewer' ? mark.text : mark.kind === 'fixed' ? 'fixed' : '', forensic: false },
      projectionMode: p.projection, policyDetails: { projection: sd.doc.contentPolicyId, ...(missing ? { missingFaces: missing } : {}) }, unitId: unitIndex as number,
    }
    const key = renderKey(fields)
    let unit = renderCache.get(key)
    if (!unit) {
      const size = sd.doc.units[unitIndex as number]!
      try {
        unit = await o.pool.render(s.sid, { docRef: sd.doc.docRef, contentPolicy: p.contentPolicy, unitIndex: unitIndex as number, scale: scale as number, projection: p.projection, mark, ...(missing ? { missingFaces: missing } : {}) }, size.w * size.h * 4 * (scale as number) ** 2 * 2)
      } catch (e) {
        if (e instanceof AdmissionRefused) return refuse(ws, reqId, 'rate-limited')
        return refuse(ws, reqId, 'policy-refused')
      }
      renderCache.put(key, unit, unit.kind === 'raster' ? unit.png.length : unit.ops.length + unit.assets.reduce((n, a) => n + (a.type === 'bitmap' ? a.rgba.length : 256), 0))
    }
    if (ws.data.phase !== 'stream' || table.get(s.sid)?.activeSocket !== (ws as unknown)) return
    renderCache.pin(key)
    try { sendUnit(ws, reqId, unitIndex as number, unit) } finally { renderCache.unpin(key) }
  }

  function sendUnit(ws: WS, reqId: number, unitIndex: number, unit: RenderedUnit) {
    if (unit.kind === 'commands') {
      let payloads: Uint8Array[] | undefined
      try { payloads = encodeAssets(unit.assets) } catch (e) { if (!(e instanceof AssetTooLarge)) throw e }
      if (payloads) {
        for (const p of payloads) sendBinary(ws, KIND.ASSET_DEFINE, reqId, p)
        if (sendBinary(ws, KIND.UNIT_COMMANDS, reqId, encodeUnit({ unitId: unitIndex, discriminator: 1, body: unit.ops }))) return
      }
      return refuse(ws, reqId, 'policy-refused')
    }
    if (unit.fallback === 'face-missing') send(ws, { reqId, type: 'diag', code: 'face-missing', unitIndex })
    if (!sendBinary(ws, KIND.UNIT_RASTER, reqId, encodeUnit({ unitId: unitIndex, discriminator: 2, body: encodeRasterBody(unit.w, unit.h, unit.png) }))) refuse(ws, reqId, 'policy-refused')
  }

  function onRenewAck(ws: WS, msg: Record<string, unknown>) {
    const sid = ws.data.sid, s = sid ? table.get(sid) : undefined
    if (!s || typeof msg.jti !== 'string' || !table.ack(s.sid, msg.jti)) return refuse(ws, msg.reqId, 'policy-refused')
    const exp = (s as Session & { pendingExp?: number }).pendingExp!
    sessionDoc.get(s.sid)!.tokenExp = exp
    scheduleRenew(ws, s.sid, exp)
  }

  function validHello(m: Record<string, unknown>): boolean {
    if (typeof m.client !== 'string' || m.client.length > 64) return false
    if (!Array.isArray(m.recVersions) || m.recVersions.length === 0 || m.recVersions.length > 16 || !m.recVersions.every((v) => Number.isInteger(v) && v >= 0 && v <= 255)) return false
    if (m.token !== undefined && (typeof m.token !== 'string' || m.token.length > 512)) return false
    if (m.resume !== undefined && (typeof m.resume !== 'object' || m.resume === null)) return false
    return true
  }

  async function onText(ws: WS, text: string) {
    const bytes = enc.encode(text).length
    if (bytes > (ws.data.phase === 'hello' ? HELLO_CAP : TEXT_CAP)) return close(ws, 'unsupported-version')
    let msg: unknown
    try { msg = JSON.parse(text) } catch { msg = undefined }
    const m = msg && typeof msg === 'object' && !Array.isArray(msg) ? msg as Record<string, unknown> : undefined
    if (ws.data.phase === 'hello') {
      if (!m || m.type !== 'hello' || !isU32(m.reqId) || !validHello(m)) return close(ws, 'unsupported-version')
      if (!(m.recVersions as number[]).includes(1)) return close(ws, 'unsupported-version')
      ws.data.phase = 'opening'
      return openSession(ws, m, m.reqId as number)
    }
    if (ws.data.sid) {
      const rate = ws.data.rates.check('message', now())
      if (rate === 'close') return close(ws, 'rate-limited')
      if (rate === 'refuse') return refuse(ws, m?.reqId, 'rate-limited')
    }
    if (!m) return refuse(ws, 0, 'policy-refused')
    // The client has no relax surface (§6.1): any attempt closes 4009.
    if (m.type === 'relax' || 'policy' in m || 'projection' in m || 'contentPolicy' in m) return close(ws, 'policy-refused')
    if (ws.data.phase === 'opening') return refuse(ws, m.reqId, 'policy-refused')
    if (m.type === 'ready') return onReady(ws, m)
    if (m.type === 'renew-ack') return onRenewAck(ws, m)
    if (m.type === 'unit') return onUnit(ws, m)
    return refuse(ws, m.reqId, 'policy-refused')
  }

  const server = {
    fetch(req: Request, srv: Server<Conn>): Response | undefined {
      const url = new URL(req.url)
      if (req.headers.get('upgrade')?.toLowerCase() !== 'websocket' || url.pathname !== (o.path ?? '/stream')) return new Response('Not Found', { status: 404 })
      const origin = req.headers.get('origin') ?? ''
      if (!o.originAllowlist.includes(origin)) return new Response('Forbidden', { status: 403 })
      if (sockets.size >= PROCESS_SOCKETS) return new Response('Busy', { status: 503 })
      const ip = srv.requestIP(req)?.address ?? 'unknown'
      const gate = ipRates.admit(ip, now())
      if (gate !== 'ok') return new Response('Too Many Requests', { status: 429 })
      const data: Conn = { ip, origin, url: req.url, headers: req.headers, phase: 'hello', rates: new SessionRates(), timers: new Set(), missingFaces: [] }
      if (srv.upgrade(req, { data })) return undefined
      ipRates.release(ip)
      return new Response('Upgrade failed', { status: 400 })
    },
    websocket: {
      maxPayloadLength: FRAGMENT,
      open(ws: WS) {
        sockets.add(ws)
        later(ws, helloMs, () => { if (ws.data.phase === 'hello') close(ws, 'ready-timeout') })
      },
      async message(ws: WS, data: string | Buffer) {
        if (ws.data.phase === 'closed') return
        if (typeof data !== 'string') {
          // No client→server binary kinds exist in Stage 1: drop before any parse.
          if (ws.data.phase === 'hello') return close(ws, 'unsupported-version')
          const rate = ws.data.rates.check('message', now())
          if (rate === 'close') close(ws, 'rate-limited')
          return
        }
        try { await onText(ws, data) } catch { close(ws, 'auth-failed') }
      },
      close(ws: WS) {
        sockets.delete(ws)
        ipRates.release(ws.data.ip)
        for (const t of ws.data.timers) clearTimeout(t)
        const sid = ws.data.sid
        ws.data.phase = 'closed'
        if (!sid) return
        const s = table.get(sid)
        if (s && s.activeSocket === (ws as unknown)) {
          table.close(sid)
          const sd = sessionDoc.get(sid)
          sessionDoc.delete(sid)
          if (sd) {
            const doc = docs.get(sd.doc.docRef)
            if (doc && --doc.refs <= 0) { docs.delete(doc.docRef); o.pool.forget(doc.docRef) }
          }
        }
      },
    },
    /** §9.3 write-off: close once with 4005. */
    revoke(sid: string): boolean {
      const s = table.get(sid)
      if (!s) return false
      close(s.activeSocket as unknown as WS, 'revoked')
      return true
    },
    audit,
    table,
    stats: () => ({ sockets: sockets.size, sessions: table.size, docs: docs.size, cache: { entries: renderCache.count, bytes: renderCache.size }, pool: o.pool.stats() }),
  }
  return server
}
