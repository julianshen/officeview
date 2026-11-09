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
import { ASSET_ALLOWANCE } from '../../src/stream/remote/reservation'
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
  /** Resume grace after an abnormal drop (1001/1006), capped by the token exp. Default 30 s. */
  detachGraceMs?: number
  /** Current layout fingerprint for a policy id (tests inject a change). */
  fingerprint?: (contentPolicyId: string) => string
  path?: string
  /** §14 document-cache ceiling (tests shrink it). */
  documentCacheBytes?: number
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
  const readyMs = o.readyMs ?? 2000, helloMs = o.helloMs ?? 5000, graceMs = o.detachGraceMs ?? 30_000
  /** sid → reaper of a detached session. */
  const reapers = new Map<string, ReturnType<typeof setTimeout>>()
  const fingerprintOf = o.fingerprint ?? layoutFingerprint
  const table = new SessionTable()
  const audit = new AuditTable()
  const ipRates = new IpRates()
  const docs = new Map<string, DocState>()          // docRef → state
  const sessionDoc = new Map<string, { doc: DocState; policy: EffectivePolicy; decision: AuthDecision; tokenExp: number }>()
  const sockets = new Set<WS>()
  /** Per-sid rate state survives resume (it belongs to the session, not the socket). */
  const sessionRates = new Map<string, SessionRates>()
  const renderCache = createRenderCache<RenderedUnit>()
  const ledger = new DocumentLedger((docRef, policy) => o.pool.dropParsed(docRef, policy === 'full' ? 'full' : 'visible-only'), o.documentCacheBytes)
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
  /** §14 per-kind encoded (`totalLen`) caps. A message over its cap is never
   * framed, so a conforming client never sees an encoded-cap rejection. */
  const ENCODED: Record<number, number> = { [KIND.MANIFEST]: 256 << 10, [KIND.ASSET_DEFINE]: 1 << 20, [KIND.UNIT_RASTER]: 8 << 20, [KIND.UNIT_COMMANDS]: 8 << 20 }
  const frame = (kind: number, reqId: number, payload: Uint8Array): Uint8Array[] | undefined => {
    const z = new Uint8Array(gzipSync(payload, { level: 6 }))
    if (z.length + 4 > ENCODED[kind]!) return undefined
    return encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: () => z, fragmentBytes: FRAGMENT })
  }
  const sendFrames = (ws: WS, messages: Uint8Array[][]) => { for (const m of messages) for (const f of m) ws.sendBinary(f) }
  const gone = (ws: WS) => ws.data.phase === 'closed'
  const owns = (ws: WS, sid: string) => !gone(ws) && table.get(sid)?.activeSocket === (ws as unknown)

  /** Document references across sessions; the last release forgets it in the pool. */
  function acquireDoc(docRef: string): void {
    const d = docs.get(docRef)
    if (d) d.refs++; else docs.set(docRef, { docRef, refs: 1 } as DocState)
  }
  function releaseDoc(docRef: string): void {
    const d = docs.get(docRef)
    if (d && --d.refs <= 0) { docs.delete(docRef); o.pool.forget(docRef) }
  }

  function markOf(p: EffectivePolicy, d: AuthDecision): MarkSpec {
    if (p.visibleMark === 'off') return { kind: 'off' }
    if (p.visibleMark === 'viewer') return { kind: 'viewer', text: d.viewerId ?? 'viewer' }
    return { kind: 'fixed' }
  }

  /** RENEW ≥ lead before exp; without an ack the old token dies at its own exp.
   * Timers belong to the owning socket (cleared on close/supersession) and
   * re-check ownership after every await. */
  function scheduleRenew(ws: WS, sid: string, exp: number) {
    later(ws, (exp - lead) * 1000 - now(), () => { void renew(ws, sid, exp).catch(() => close(ws, 'session-expired')) })
  }
  async function renew(ws: WS, sid: string, exp: number) {
    const sd = sessionDoc.get(sid)
    if (!sd || !owns(ws, sid)) return
    const next = await issueToken(o.keys.kToken, { docRef: sd.doc.docRef, sid, ttlSec: ttl, now: nowSec() })
    if (!owns(ws, sid)) return
    const s = table.get(sid)!
    table.beginRotation(sid, next.jti)
    ;(s as Session & { pendingExp?: number }).pendingExp = next.exp
    send(ws, { reqId: 0, type: 'renew', token: next.token, jti: next.jti, exp: next.exp })
    later(ws, exp * 1000 - now(), () => { if (owns(ws, sid) && table.get(sid)!.pendingJti === next.jti) close(ws, 'session-expired') })
  }

  async function openSession(ws: WS, hello: Record<string, unknown>, reqId: number) {
    const req: AuthRequest = { headers: ws.data.headers, ip: ws.data.ip, url: new URL(ws.data.url), ...(typeof hello.token === 'string' ? { token: hello.token } : {}) }
    const resume = hello.resume as { sid?: unknown; resumeProof?: unknown } | undefined
    if (resume !== undefined) return resumeSession(ws, req, resume, reqId)
    const decision = await o.authorize(req)
    if (gone(ws)) return
    if (!decision) return close(ws, 'auth-failed')
    const resolved = await o.resolveDocument(req, decision)
    if (gone(ws)) return
    if (!resolved) return close(ws, 'auth-failed')
    if (!admit(resolved.bytes).ok) return close(ws, 'unsupported-format')
    const policy = effectivePolicy(decision)
    const docRef = await docRefOf(o.keys.kDoc, resolved.bytes)
    if (gone(ws)) return
    // Charge the document BEFORE parsing; one that cannot fit is refused at admission.
    if (!ledger.touch(docRef, policy.contentPolicy, resolved.bytes.length)) return close(ws, 'unsupported-format')
    acquireDoc(docRef)
    let committed = false
    try {
      let opened
      try { opened = await o.pool.open(docRef, resolved.bytes, policy.contentPolicy) } catch (e) {
        return close(ws, (e as RenderFailed).refusal === 'unsupported-format' ? 'unsupported-format' : 'auth-failed')
      }
      if (gone(ws)) return
      const sid = newSid()
      const token = await issueToken(o.keys.kToken, { docRef, sid, ttlSec: ttl, now: nowSec() })
      if (gone(ws)) return
      // Commit synchronously: nothing below awaits, so no close can interleave.
      const s = table.create({ sid, docRef, ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), role: policy.role, effectivePolicy: policy, activeSocket: ws as never, jti: token.jti, startedAt: nowSec() })
      sessionDoc.set(sid, { doc: { docRef, name: resolved.name ?? 'document', bytes: resolved.bytes, units: opened.units, fingerprint: opened.layoutFingerprint, contentPolicyId: opened.contentPolicyId, format: opened.format, refs: 0 }, policy, decision, tokenExp: token.exp })
      sessionRates.set(sid, ws.data.rates)
      audit.record(sid, { ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), ip: ws.data.ip, startedAt: s.startedAt, keyGeneration: 1, policyId: opened.contentPolicyId, markSettings: { visible: policy.visibleMark, forensic: false, forensicRequested: policy.forensicRequested }, relaxed: policy.relaxed, forensic: null, forensicLayers: [] })
      ws.data.sid = sid
      committed = true
      sendManifest(ws, reqId, sid, token.token)
      scheduleRenew(ws, sid, token.exp)
    } finally { if (!committed) releaseDoc(docRef) }
  }

  /** §6.2 rule 7. Resume ALWAYS re-runs `authorize` and re-evaluates the
   * effective policy; the proof is the live session token, or a proof the hook
   * itself grants (`resumeGranted`). Validation and supersession are bound to
   * the jti generation observed before any await, so of two racing resumes
   * only one can commit. */
  async function resumeSession(ws: WS, req: AuthRequest, resume: { sid?: unknown; resumeProof?: unknown }, reqId: number) {
    if (typeof resume.sid !== 'string' || typeof resume.resumeProof !== 'string') return close(ws, 'auth-failed')
    const s = table.get(resume.sid), sd = sessionDoc.get(resume.sid)
    if (!s || !sd) return close(ws, 'auth-failed')
    const generation = s.jti
    const v = await verifyToken(o.keys.kToken, resume.resumeProof, { now: nowSec(), docRef: s.docRef, sid: s.sid })
    if (gone(ws)) return
    const tokenProof = v.ok && ['current', 'rotating'].includes(table.acceptJti(s.sid, v.claims.jti))
    const d = await o.authorize({ ...req, resume: { sid: s.sid }, ...(tokenProof ? {} : { token: resume.resumeProof }) })
    if (gone(ws)) return
    const ok = !!d && (tokenProof || d.resumeGranted === true) && d.viewerId === s.viewerId && JSON.stringify(effectivePolicy(d)) === JSON.stringify(s.effectivePolicy)
    if (!ok || table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    const token = await issueToken(o.keys.kToken, { docRef: s.docRef, sid: s.sid, ttlSec: ttl, now: nowSec() })
    if (gone(ws)) return
    if (table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    const old = s.activeSocket as unknown as WS | undefined
    if (old && old !== ws && old.data) { old.data.phase = 'closed'; for (const t of old.data.timers) clearTimeout(t) }
    const reaper = reapers.get(s.sid); if (reaper) { clearTimeout(reaper); reapers.delete(s.sid) }
    table.supersede(s.sid, ws as never, token.jti)
    sd.tokenExp = token.exp
    sd.decision = d!
    ws.data.sid = s.sid
    ws.data.rates = sessionRates.get(s.sid) ?? ws.data.rates
    sendManifest(ws, reqId, s.sid, token.token)
    scheduleRenew(ws, s.sid, token.exp)
  }

  function teardown(sid: string) {
    const r = reapers.get(sid); if (r) { clearTimeout(r); reapers.delete(sid) }
    table.close(sid)
    sessionRates.delete(sid)
    const sd = sessionDoc.get(sid)
    sessionDoc.delete(sid)
    if (sd) releaseDoc(sd.doc.docRef)
  }
  /** Keep the session (and its document charge) for a resume until the grace
   * or the newest token's exp, whichever comes first. */
  function detach(sid: string) {
    table.detach(sid)
    const sd = sessionDoc.get(sid)!, s = table.get(sid)! as Session & { pendingExp?: number }
    const exp = Math.max(sd.tokenExp, s.pendingExp ?? 0)
    reapers.set(sid, setTimeout(() => { reapers.delete(sid); if (table.get(sid)?.status === 'detached') teardown(sid) }, Math.max(0, Math.min(graceMs, exp * 1000 - now()))))
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
    sendFrames(ws, [frame(KIND.MANIFEST, reqId, encodeManifest(m))!])
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
    if (!owns(ws, s.sid)) return
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
    // Every request re-charges its parse: an evicted document re-enters the ledger on reparse.
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    const p = sd.policy, mark = markOf(p, sd.decision)
    const missing = p.projection === 'commands' ? [...ws.data.missingFaces].sort() : undefined
    const job = { docRef: sd.doc.docRef, contentPolicy: p.contentPolicy, unitIndex: unitIndex as number, scale: scale as number, mark }
    const estimate = (() => { const size = sd.doc.units[unitIndex as number]!; return size.w * size.h * 4 * (scale as number) ** 2 * 2 })()
    const fields = (projectionMode: 'final-raster' | 'commands'): RenderKeyFields => ({
      docRef: sd.doc.docRef, contentPolicyId: sd.doc.contentPolicyId, layoutFingerprint: sd.doc.fingerprint, recorderVersion: 1, scaleBucket: scale as number,
      markSettings: { visible: mark.kind, text: mark.kind === 'viewer' ? mark.text : mark.kind === 'fixed' ? 'fixed' : '', forensic: false },
      projectionMode, policyDetails: { projection: sd.doc.contentPolicyId, ...(projectionMode === 'commands' && missing ? { missingFaces: missing } : {}) }, unitId: unitIndex as number,
    })
    const obtain = async (projection: 'final-raster' | 'commands'): Promise<RenderedUnit | RefusalCode> => {
      const key = renderKey(fields(projection))
      const hit = renderCache.get(key)
      if (hit) return hit
      try {
        const unit = await o.pool.render(s.sid, { ...job, projection, ...(projection === 'commands' && missing ? { missingFaces: missing } : {}) }, estimate)
        renderCache.put(key, unit, unit.kind === 'raster' ? unit.png.length : unit.ops.length + unit.assets.reduce((n, a) => n + (a.type === 'bitmap' ? a.rgba.length : 256), 0))
        return unit
      } catch (e) { return e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused' }
    }
    let unit = await obtain(p.projection)
    if (!owns(ws, s.sid) || ws.data.phase !== 'stream') return
    if (typeof unit === 'string') return refuse(ws, reqId, unit)
    let messages = prepareUnit(reqId, unitIndex as number, unit)
    if (!messages && unit.kind === 'commands') {
      // Assets too large for ASSET_DEFINE's caps: ship the unit as a raster (Task 7).
      unit = await obtain('final-raster')
      if (!owns(ws, s.sid) || ws.data.phase !== 'stream') return
      if (typeof unit === 'string') return refuse(ws, reqId, unit)
      messages = prepareUnit(reqId, unitIndex as number, unit)
    }
    if (!messages) return refuse(ws, reqId, 'policy-refused')
    if (unit.kind === 'raster' && unit.fallback === 'face-missing') send(ws, { reqId, type: 'diag', code: 'face-missing', unitIndex })
    sendFrames(ws, messages)
  }

  /** Frames a unit's messages under every per-kind cap, or undefined (nothing sent). */
  function prepareUnit(reqId: number, unitIndex: number, unit: RenderedUnit): Uint8Array[][] | undefined {
    if (unit.kind === 'commands') {
      let payloads: Uint8Array[]
      try { payloads = encodeAssets(unit.assets) } catch (e) { if (e instanceof AssetTooLarge) return undefined; throw e }
      // §14: the client admits at most ASSET_ALLOWANCE asset bytes per request; more would only fail there.
      if (payloads.reduce((n, p) => n + p.length, 0) > ASSET_ALLOWANCE) return undefined
      const out: Uint8Array[][] = []
      for (const p of payloads) { const f = frame(KIND.ASSET_DEFINE, reqId, p); if (!f) return undefined; out.push(f) }
      const f = frame(KIND.UNIT_COMMANDS, reqId, encodeUnit({ unitId: unitIndex, discriminator: 1, body: unit.ops }))
      return f ? [...out, f] : undefined
    }
    const f = frame(KIND.UNIT_RASTER, reqId, encodeUnit({ unitId: unitIndex, discriminator: 2, body: encodeRasterBody(unit.w, unit.h, unit.png) }))
    return f ? [f] : undefined
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
      close(ws: WS, code: number) {
        sockets.delete(ws)
        ipRates.release(ws.data.ip)
        for (const t of ws.data.timers) clearTimeout(t)
        const sid = ws.data.sid
        // A server-initiated close set the phase first; only a peer drop reaches here open.
        const serverClosed = ws.data.phase === 'closed'
        ws.data.phase = 'closed'
        if (!sid) return
        const s = table.get(sid)
        if (!s || s.activeSocket !== (ws as unknown)) return
        if (!serverClosed && (code === 1001 || code === 1006) && (s.status === 'manifest' || s.status === 'stream')) return detach(sid)
        teardown(sid)
      },
    },
    /** §9.3 write-off: close once with 4005. */
    revoke(sid: string): boolean {
      const s = table.get(sid)
      if (!s) return false
      if (!s.activeSocket) { teardown(sid); return true }
      close(s.activeSocket as unknown as WS, 'revoked')
      return true
    },
    audit,
    table,
    stats: () => ({ sockets: sockets.size, sessions: table.size, docs: docs.size, cache: { entries: renderCache.count, bytes: renderCache.size }, pool: o.pool.stats() }),
  }
  return server
}
