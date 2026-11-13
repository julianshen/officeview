/**
 * The only endpoint (§9.1): WS upgrade with an origin allowlist, the §6.2
 * state machine (HELLO → MANIFEST → READY → STREAM, RENEW, RESUME), per-op
 * token + capability checks, §14 rate gates, and the §6.2 rule-8 close-code
 * registry. No HTTP document route exists: every non-upgrade request is 404.
 */
import { createHash } from 'node:crypto'
import { gzipSync } from 'node:zlib'
import type { Server, ServerWebSocket } from 'bun'
import { encodeMessage } from '../../src/stream/frame'
import { encodeUnit } from '../../src/stream/unit'
import { encodeRasterBody } from '../../src/stream/raster'
import { AssetTooLarge, encodeAssets } from '../../src/stream/asset-wire'
import { ASSET_ALLOWANCE } from '../../src/stream/remote/reservation'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { AuditTable, ForensicExhausted, msgHex, type Issued } from './audit'
import { CODEC_ID } from './mark/params'
import { encodeEstimate } from './mark/mark-raster'
import { deriveBasisKey } from './mark/basis'
import { LeaseFull, ScratchLease } from './scratch'
import { createRenderCache, DocumentLedger, renderKey, type RenderKeyFields } from './cache'
import { loadPack, requiredFacesOf } from './fonts'
import { effectivePolicy, forensicNotices, resumeInvariant, type AuthDecision, type AuthRequest, type Authorize, type EffectivePolicy, type ResolveDocument } from './hooks'
import { docRefOf, type Keys } from './keys'
import { admit } from './policy'
import { AdmissionRefused, RenderFailed, type Pool } from './pool'
import { encodePrintBody, planPrint, PRINT_ENCODED_CAP, PRINT_SEND_HIGH_WATER, printArgs, printSizeOk, runPrint, type PrintJob, type PrintSlot } from './print'
import { IpRates, SessionRates } from './rate'
import { layoutFingerprint, SCALE_BUCKETS, type MarkSpec, type RenderedUnit, type UnitSize } from './render'
import { SessionTable, type Session } from './sessionTable'
import { a11yArgs, rangeArg, runA11y, runCopy, runSearch, runSelect, searchArgs, type SearchRun, type TextSlot } from './text/ops'
import { issueToken, newSid, verifyToken } from './tokens'

/** §6.2 rule 8 — the complete registry. 4007 is client-only and never sent here. */
export const CLOSE = {
  'unsupported-version': 4000, 'unsupported-format': 4001, 'ready-timeout': 4002, 'auth-failed': 4003,
  'session-expired': 4004, 'revoked': 4005, 'layout-changed': 4006, 'rate-limited': 4008, 'policy-refused': 4009,
} as const
export type CloseName = keyof typeof CLOSE
export type RefusalCode = 'rate-limited' | 'policy-refused' | 'session-rotated' | 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted' | 'print-too-large' | 'print-quota-exhausted'

const KIND = { MANIFEST: 1, ASSET_DEFINE: 2, UNIT_RASTER: 3, UNIT_COMMANDS: 4, PRINT_RASTER: 8 } as const
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
  /** Audit table over its durable store (default: in-memory). */
  audit?: AuditTable
  /** TEST-ONLY: wedge each Layer-1 encode (forwarded only by a testHooks pool). */
  encodeHangMs?: number
  /** TEST-ONLY: replaces the print job planner (e.g. to force a too-large refusal). */
  planPrint?: typeof planPrint
  /** TEST-ONLY: interposes on binary sends and the socket's buffered amount (backpressure tests). */
  transport?: { sendBinary(ws: unknown, frame: Uint8Array): number; bufferedAmount(ws: unknown): number }
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
  /** Print jobs waiting for this socket to drain. */
  drainWaiters: Set<() => void>
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
  const audit = o.audit ?? new AuditTable()
  const ipRates = new IpRates()
  const docs = new Map<string, DocState>()          // docRef → state
  /** Per-sid session entry. Its text state (in-flight search/select, the copy
   * queue and the quota counter `copiedParagraphs`) survives resume and is
   * gone on teardown. So does its print state (`printedPages`, the running job). */
  const sessionDoc = new Map<string, TextSlot & PrintSlot & { doc: DocState; policy: EffectivePolicy; decision: AuthDecision; tokenExp: number; /** Layer-1 issuance, fixed for the sid's life (a resume keeps it). */ forensic?: Issued; /** The sid's in-flight search (R6: one per sid); survives resume. */ search?: SearchRun }>()
  const sockets = new Set<WS>()
  /** Per-sid rate state survives resume (it belongs to the session, not the socket). */
  const sessionRates = new Map<string, SessionRates>()
  const renderCache = createRenderCache<RenderedUnit>()
  /** sid → §14 scratch lease (≤ 2 Layer-1 encodes in flight). */
  const leases = new Map<string, ScratchLease>()
  const forensicStats = { marked: 0, unmarked: 0, carriers: [] as number[], leased: 0 }
  const leaseFor = (sid: string) => { let l = leases.get(sid); if (!l) leases.set(sid, l = new ScratchLease(2, o.pool.admission.caps.perSessionQueue, (d) => { forensicStats.leased += d })); return l }
  const ledger = new DocumentLedger((docRef, policy) => o.pool.dropParsed(docRef, policy === 'full' ? 'full' : 'visible-only'), o.documentCacheBytes)
  const pack = loadPack()

  const send = (ws: WS, msg: object) => { if (ws.data.phase !== 'closed') ws.send(JSON.stringify(msg)) }
  /** `limit`/`requested`: the cap a refusal states (copy caps, Task 5). */
  const refuse = (ws: WS, reqId: unknown, code: RefusalCode, nums: { limit?: number; requested?: number } = {}) => send(ws, { reqId: isU32(reqId) ? reqId : 0, type: 'refused', code, ...nums })
  const close = (ws: WS, name: CloseName | 'superseded') => {
    if (ws.data.phase === 'closed') return
    const code = name === 'superseded' ? 4409 : CLOSE[name]
    ws.data.phase = 'closed'
    wakeDrain(ws)
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
  const ENCODED: Record<number, number> = { [KIND.MANIFEST]: 256 << 10, [KIND.ASSET_DEFINE]: 1 << 20, [KIND.UNIT_RASTER]: 8 << 20, [KIND.UNIT_COMMANDS]: 8 << 20, [KIND.PRINT_RASTER]: PRINT_ENCODED_CAP }
  const frame = (kind: number, reqId: number, payload: Uint8Array): Uint8Array[] | undefined => {
    const z = new Uint8Array(gzipSync(payload, { level: 6 }))
    if (z.length + 4 > ENCODED[kind]!) return undefined
    return encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: () => z, fragmentBytes: FRAGMENT })
  }
  const sendBinary = (ws: WS, f: Uint8Array) => o.transport ? o.transport.sendBinary(ws, f) : ws.sendBinary(f)
  const bufferedOf = (ws: WS) => o.transport ? o.transport.bufferedAmount(ws) : ws.getBufferedAmount()
  /** Sends every fragment in order; false (and nothing more sent) once the transport drops one (Bun: 0).
   * -1 (backpressure) means queued. The unit path ignores the result (pre-existing behaviour). */
  const sendFrames = (ws: WS, messages: Uint8Array[][]): boolean => {
    for (const m of messages) for (const f of m) if (sendBinary(ws, f) === 0) return false
    return true
  }
  const wakeDrain = (ws: WS) => { const w = [...ws.data.drainWaiters]; ws.data.drainWaiters.clear(); for (const f of w) f() }
  /** Resolves once `ws` holds at most PRINT_SEND_HIGH_WATER unsent bytes, or `alive` fails.
   * Woken by `drain`, the socket's close, or the job's `wake` (cancel, teardown, resume);
   * a slow re-check guards a missed wake. */
  async function belowHighWater(ws: WS, job: PrintJob, alive: () => boolean) {
    while (alive() && bufferedOf(ws) > PRINT_SEND_HIGH_WATER) {
      await new Promise<void>((resolve) => {
        const wake = () => { clearTimeout(t); ws.data.drainWaiters.delete(wake); if (job.wake === wake) job.wake = undefined; resolve() }
        const t = setTimeout(wake, 5_000)
        ws.data.drainWaiters.add(wake); job.wake = wake
      })
    }
  }
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
      let forensic: Issued | undefined
      if (policy.forensicMark) {
        try { forensic = audit.issue(o.keys.kMsg, o.keys.keyGeneration) } catch (e) {
          if (e instanceof ForensicExhausted) return close(ws, 'policy-refused')
          throw e
        }
      }
      const sid = newSid()
      const token = await issueToken(o.keys.kToken, { docRef, sid, ttlSec: ttl, now: nowSec() })
      if (gone(ws)) return
      // Commit synchronously: nothing below awaits, so no close can interleave.
      // The issuance row is durable before anything is committed or sent; a failed write refuses the session.
      const startedAt = nowSec()
      try {
        audit.record(sid, { ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), ip: ws.data.ip, startedAt, keyGeneration: o.keys.keyGeneration, policyId: opened.contentPolicyId, markSettings: { visible: policy.visibleMark, forensicRequested: policy.forensicRequested }, relaxed: policy.relaxed, services: { copy: policy.copy, textA11y: policy.textA11y, print: policy.print }, forensic: forensic ? { codec: CODEC_ID, token24: forensic.token24, salt8: forensic.salt8, msg40: msgHex(forensic.msg40) } : null, docRef, contentPolicy: policy.contentPolicy, visibleMark: markOf(policy, decision) })
      } catch { return close(ws, 'policy-refused') }
      table.create({ sid, docRef, ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), role: policy.role, effectivePolicy: policy, activeSocket: ws as never, jti: token.jti, startedAt })
      sessionDoc.set(sid, { doc: { docRef, name: resolved.name ?? 'document', bytes: resolved.bytes, units: opened.units, fingerprint: opened.layoutFingerprint, contentPolicyId: opened.contentPolicyId, format: opened.format, refs: 0 }, policy, decision, tokenExp: token.exp, copiedParagraphs: 0, printedPages: 0, ...(forensic ? { forensic } : {}) })
      sessionRates.set(sid, ws.data.rates)
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
    // R3: copy, textA11y and the print caps (Stage 3b P4) may change on resume; every other field is exact-match.
    const policy = d ? effectivePolicy(d) : undefined
    const ok = !!d && !!policy && (tokenProof || d.resumeGranted === true) && d.viewerId === s.viewerId && JSON.stringify(resumeInvariant(policy)) === JSON.stringify(resumeInvariant(s.effectivePolicy))
    if (!ok || table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    const token = await issueToken(o.keys.kToken, { docRef: s.docRef, sid: s.sid, ttlSec: ttl, now: nowSec() })
    if (gone(ws)) return
    if (table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    // A change of the mutable text services is durable before it takes effect
    // (the open row stays write-once); a failed write refuses the resume.
    if (JSON.stringify([policy!.copy, policy!.textA11y, policy!.print]) !== JSON.stringify([sd.policy.copy, sd.policy.textA11y, sd.policy.print])) {
      try { audit.recordServices(s.sid, { copy: policy!.copy, textA11y: policy!.textA11y, print: policy!.print }) } catch { return close(ws, 'policy-refused') }
    }
    const old = s.activeSocket as unknown as WS | undefined
    if (old && old !== ws && old.data) { old.data.phase = 'closed'; for (const t of old.data.timers) clearTimeout(t); wakeDrain(old) }
    // A job on the old socket cannot continue on the new one.
    if (sd.print) { sd.print.stopped = true; sd.print.wake?.() }
    const reaper = reapers.get(s.sid); if (reaper) { clearTimeout(reaper); reapers.delete(s.sid) }
    table.supersede(s.sid, ws as never, token.jti)
    sd.tokenExp = token.exp
    sd.decision = d!
    sd.policy = policy!
    table.setPolicy(s.sid, policy!)
    ws.data.sid = s.sid
    ws.data.rates = sessionRates.get(s.sid) ?? ws.data.rates
    sendManifest(ws, reqId, s.sid, token.token)
    scheduleRenew(ws, s.sid, token.exp)
  }

  function teardown(sid: string) {
    const r = reapers.get(sid); if (r) { clearTimeout(r); reapers.delete(sid) }
    table.close(sid)
    sessionRates.delete(sid)
    leases.get(sid)?.close(); leases.delete(sid)
    const sd = sessionDoc.get(sid)
    sessionDoc.delete(sid)
    if (sd?.search) { sd.search.stopped = true; sd.search = undefined }
    if (sd?.select) { sd.select.stopped = true; sd.select = undefined }
    if (sd?.print) { sd.print.stopped = true; sd.print.wake?.() }
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
      // Presentation, notes and links land later in Stage 3.
      can: { search: true, copy: p.copy !== false, print: p.print !== false, textA11y: p.textA11y !== 'off', present: false, notes: false, links: false },
      caps: { copy: p.copy === false ? null : { ...p.copy }, a11y: p.textA11y, print: p.print === false ? null : { ...p.print } },
      recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [...SCALE_BUCKETS], requiredFaces: p.projection === 'commands' ? requiredFacesOf(pack.faces.map((f) => f.id), pack) : [], projection: p.projection },
      limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: FRAGMENT, tilePx: 512 },
      notices: [...forensicNotices(p), ...(p.relaxed ? ['relaxed-session'] : [])],
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
    const p = sd.policy
    const missing = p.projection === 'commands' ? [...ws.data.missingFaces].sort() : undefined
    const alive = () => owns(ws, s.sid) && ws.data.phase === 'stream'
    const ui = unitIndex as number, sc = scale as number
    const first = await obtainUnit(s.sid, sd, ui, sc, p.projection, missing)
    if (!alive()) return
    if (typeof first === 'string') return refuse(ws, reqId, first)
    let unit: RenderedUnit = first
    let messages = unit.kind === 'commands' ? prepareUnit(reqId, ui, unit) : undefined
    if (!messages && unit.kind === 'commands') {
      // Assets too large for ASSET_DEFINE's caps: ship the unit as a raster (Task 7).
      const raster = await obtainUnit(s.sid, sd, ui, sc, 'final-raster', missing)
      if (!alive()) return
      if (typeof raster === 'string') return refuse(ws, reqId, raster)
      unit = raster
    }
    let unmarked: string | undefined
    if (unit.kind === 'raster') {
      const d = await deliverRaster(alive, s.sid, sd, ui, sc, unit, (u) => prepareUnit(reqId, ui, u))
      if (d === undefined) return
      if (typeof d === 'string') return refuse(ws, reqId, d)
      messages = d.messages; unmarked = d.unmarked
    }
    if (!messages) return refuse(ws, reqId, 'policy-refused')
    if (unit.kind === 'raster' && unit.fallback === 'face-missing') send(ws, { reqId, type: 'diag', code: 'face-missing', unitIndex })
    if (unmarked) send(ws, { reqId, type: 'diag', code: 'forensic-unmarked', unitIndex, reason: unmarked })
    sendFrames(ws, messages)
  }

  /** The canonical unit for a session: the render cache, else the pool (then cached). */
  async function obtainUnit(sid: string, sd: { doc: DocState; policy: EffectivePolicy; decision: AuthDecision }, unitIndex: number, scale: number, projection: 'final-raster' | 'commands', missing?: string[]): Promise<RenderedUnit | RefusalCode> {
    const p = sd.policy, mark = markOf(p, sd.decision)
    const key = renderKey(keyFields(sd.doc, mark, unitIndex, scale, projection, missing))
    const hit = renderCache.get(key)
    if (hit) return hit
    const size = sd.doc.units[unitIndex]!, estimate = size.w * size.h * 4 * scale ** 2 * 2
    try {
      const unit = await o.pool.render(sid, { docRef: sd.doc.docRef, contentPolicy: p.contentPolicy, unitIndex, scale, mark, projection, ...(projection === 'commands' && missing ? { missingFaces: missing } : {}) }, estimate)
      renderCache.put(key, unit, unit.kind === 'raster' ? unit.png.length : unit.ops.length + unit.assets.reduce((n, a) => n + (a.type === 'bitmap' ? a.rgba.length : 256), 0))
      return unit
    } catch (e) { return e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused' }
  }

  /** The shared per-unit tail (unit and print): Layer-1 over the canonical
   * raster, frame it with `pack`, then the durable delivery record — in that
   * order, so a unit over its frame cap leaves no delivery row, and nothing
   * goes out unrecorded. `undefined`: the caller died meanwhile (send nothing). */
  async function deliverRaster(alive: () => boolean, sid: string, sd: { doc: DocState; forensic?: Issued }, unitIndex: number, scale: number, unit: RenderedUnit & { kind: 'raster' }, pack: (u: RenderedUnit & { kind: 'raster' }) => Uint8Array[][] | undefined): Promise<{ messages: Uint8Array[][]; unmarked?: string } | RefusalCode | undefined> {
    // Every raster this session receives — fallbacks included — carries its Layer-1 mark.
    let unmarked: string | undefined
    const canonical = sd.forensic ? unit.png : undefined
    if (sd.forensic) {
      const m = await markUnit(sid, sd, unitIndex, scale, unit)
      if (!alive()) return undefined
      if (typeof m === 'string') return m
      unit = m.unit; unmarked = m.unmarked
    }
    const messages = pack(unit)
    if (!messages) return 'policy-refused'
    // Durable before anything about this unit goes out (never a delivered mark without a record); a failed write refuses it.
    if (canonical) {
      try { audit.recordDelivery({ sid, unitIndex, scale, status: unmarked ? 'unmarked' : 'l1', ...(unmarked ? { reason: unmarked } : {}), refSha256: createHash('sha256').update(canonical).digest('hex') }) } catch { return 'policy-refused' }
    }
    return { messages, ...(unmarked ? { unmarked } : {}) }
  }

  /** The onUnit preamble for a text op: stream phase, reqId, token, ownership
   * after the await, jti. Answers (refuse/close) and returns undefined on failure. */
  /** `cancelled`: checked after the token await; true ⇒ nothing is sent (a print cancelled while it was validated). */
  async function textOp(ws: WS, msg: Record<string, unknown>, cancelled?: () => boolean) {
    const reqId = msg.reqId
    const sid = ws.data.sid, s = sid ? table.get(sid) : undefined, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !s || !sd || !isU32(reqId) || reqId === 0) return void refuse(ws, reqId, 'policy-refused')
    const v = await verifyToken(o.keys.kToken, msg.token, { now: nowSec(), docRef: sd.doc.docRef, sid: s.sid })
    if (!owns(ws, s.sid) || cancelled?.()) return
    if (!v.ok) return void (v.reason === 'exp' ? close(ws, 'session-expired') : refuse(ws, reqId, 'policy-refused'))
    const jti = table.acceptJti(s.sid, v.claims.jti)
    if (jti === 'rotated') return void refuse(ws, reqId, 'session-rotated')
    if (jti === 'unknown') return void refuse(ws, reqId, 'policy-refused')
    return { reqId, s, sd }
  }
  /** The issuing socket still owns the sid, in `stream`, with the same entry. */
  const textAlive = (ws: WS, sid: string, sd: unknown) => () => owns(ws, sid) && ws.data.phase === 'stream' && sessionDoc.get(sid) === sd

  /** §4.1 search: the onUnit checks, then R6 arguments and the R5 'text'
   * rate; `runSearch` walks the units and owns supersession. Between batches
   * the run re-checks ownership, the stream phase and the session entry. */
  async function onSearch(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    const args = searchArgs(msg)
    if (!args) return refuse(ws, reqId, 'policy-refused')
    // R5: a 'text' overflow is a refusal only (never a breach toward 4008).
    if (ws.data.rates.check('text', now()) !== 'ok') return refuse(ws, reqId, 'rate-limited')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    await runSearch({
      slot: sd, reqId, ...args, unitCount: sd.doc.units.length,
      batch: (job) => {
        // A long walk can outlive the parse: every job re-charges it.
        if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return Promise.reject(new Error('document cache full'))
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job })
      },
      alive: textAlive(ws, s.sid, sd),
      send: (m) => send(ws, m),
      refuse: (code) => refuse(ws, reqId, code),
    })
  }

  /** §4.1 select (rects of a range): requires `can.copy`; charged only the
   * message budget (R5). A newer select on the sid supersedes this one. */
  async function onSelect(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    if (sd.policy.copy === false) return refuse(ws, reqId, 'text-unavailable')
    const range = rangeArg(msg.range, sd.doc.units.length)
    if (!range) return refuse(ws, reqId, 'policy-refused')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    await runSelect({
      slot: sd, reqId, range,
      job: (job) => o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job }),
      alive: textAlive(ws, s.sid, sd),
      send: (m) => send(ws, m),
      refuse: (code) => refuse(ws, reqId, code),
    })
  }

  /** §4.1 copy: the copy caps, the session quota and the audit record
   * (`runCopy`); charged the R5 'text' rate. Copies on one sid are serial. */
  async function onCopy(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    if (sd.policy.copy === false) return refuse(ws, reqId, 'text-unavailable')
    const range = rangeArg(msg.range, sd.doc.units.length)
    if (!range) return refuse(ws, reqId, 'policy-refused')
    if (ws.data.rates.check('text', now()) !== 'ok') return refuse(ws, reqId, 'rate-limited')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    await runCopy({
      slot: sd, reqId, range, unitCount: sd.doc.units.length,
      caps: () => sd.policy.copy,
      job: (job) => {
        // A queued copy can outlive the parse: charge it when it runs.
        if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return Promise.reject(new Error('document cache full'))
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job })
      },
      record: (r) => audit.recordCopy({ sid: s.sid, kind: 'copy', ...r }),
      alive: textAlive(ws, s.sid, sd),
      send: (m) => send(ws, m),
      refuse: (code, nums) => refuse(ws, reqId, code, nums),
    })
  }

  /** §4.1 a11y text: charged the R5 'text' rate (before the 'off' ⇒ text-unavailable refusal);
   * `runA11y` applies the mode, the shared session quota ('focused' k) and
   * the audit record, on the sid's copy queue. */
  async function onA11y(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    const args = a11yArgs(msg, sd.doc.units.length)
    if (!args) return refuse(ws, reqId, 'policy-refused')
    // Every valid authenticated a11y request is charged, 'off' sessions included.
    if (ws.data.rates.check('text', now()) !== 'ok') return refuse(ws, reqId, 'rate-limited')
    if (sd.policy.textA11y === 'off') return refuse(ws, reqId, 'text-unavailable')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    await runA11y({
      slot: sd, reqId, ...args,
      mode: () => sd.policy.textA11y,
      caps: () => sd.policy.copy,
      job: (job) => {
        if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return Promise.reject(new Error('document cache full'))
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job })
      },
      record: (r) => audit.recordCopy({ sid: s.sid, kind: 'a11y', paragraphs: r.paragraphs, fromUnit: args.unit, toUnit: args.unit }),
      alive: textAlive(ws, s.sid, sd),
      send: (m) => send(ws, m),
      refuse: (code, nums) => refuse(ws, reqId, code, nums),
    })
  }

  /** §4.2 print (Stage 3b P3): the whole job is checked, rate-charged, capped,
   * audited and charged before `print-start`; everything after the token
   * await up to `print-start` is synchronous, so two prints cannot both pass
   * the running-job and quota checks. Pages are always final rasters, through
   * the same cache → pool → Layer-1 → delivery-record step as `unit`.
   *
   * The job takes the sid's one print slot synchronously, before the token
   * await, when the slot is free; a `print` arriving while another is pending
   * or running does not, and is refused `rate-limited` after validation (as
   * before), even if the slot has freed meanwhile. A `print-cancel` for it during validation marks it stopped; it
   * then ends silently after the await: nothing sent, rated, audited or
   * charged, and the slot freed. Only the one pending job is held (no
   * record of unknown reqIds). */
  async function onPrint(ws: WS, msg: Record<string, unknown>) {
    const sid0 = ws.data.sid, sd0 = sid0 ? sessionDoc.get(sid0) : undefined
    // The one-job rule is decided at arrival: a print that finds the slot taken
    // is refused rate-limited even if the slot frees during its validation.
    const busy = !!sd0?.print
    let job: PrintJob | undefined
    if (ws.data.phase === 'stream' && sd0 && !busy && sid0 && owns(ws, sid0) && isU32(msg.reqId) && msg.reqId !== 0) sd0.print = job = { reqId: msg.reqId, stopped: false }
    try { await printOp(ws, msg, job, busy) } finally { if (job && sd0 && sd0.print === job) sd0.print = undefined }
  }
  async function printOp(ws: WS, msg: Record<string, unknown>, pending: PrintJob | undefined, busy: boolean) {
    const a = await textOp(ws, msg, () => !!pending?.stopped)
    if (!a) return
    const { reqId, s, sd } = a
    const args = printArgs(msg, sd.doc.units.length)
    if (!args) return refuse(ws, reqId, 'policy-refused')
    // A 'print' overflow is a refusal only (never a breach toward 4008).
    if (ws.data.rates.check('print', now()) !== 'ok') return refuse(ws, reqId, 'rate-limited')
    const caps = sd.policy.print
    if (caps === false) return refuse(ws, reqId, 'policy-refused')
    if (busy || (sd.print && sd.print !== pending)) return refuse(ws, reqId, 'rate-limited')
    const { from, to } = args, n = to - from + 1
    const plan = (o.planPrint ?? planPrint)(sd.doc.units, from, to)
    if ('refused' in plan) return refuse(ws, reqId, 'print-too-large')
    if (caps.maxPages !== null && n > caps.maxPages) return refuse(ws, reqId, 'print-too-large', { limit: caps.maxPages, requested: n })
    if (caps.maxSessionPages !== null && sd.printedPages + n > caps.maxSessionPages) return refuse(ws, reqId, 'print-quota-exhausted', { limit: caps.maxSessionPages, requested: n })
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    const { scale } = plan
    try { audit.recordPrint({ sid: s.sid, pages: n, fromUnit: from, toUnit: to, scale }) } catch { return refuse(ws, reqId, 'policy-refused') }
    // P3: charged now; a job that fails later stays charged.
    sd.printedPages += n
    const job: PrintJob = pending ?? { reqId, stopped: false }
    sd.print = job
    send(ws, { reqId, type: 'print-start', pages: n, scale })
    const alive = () => !job.stopped && owns(ws, s.sid) && ws.data.phase === 'stream' && sessionDoc.get(s.sid) === sd
    try {
      await runPrint({
        job, from, to, alive,
        ready: () => belowHighWater(ws, job, alive),
        page: async (unitIndex) => {
          // A long job can outlive the parse: every page re-charges it.
          if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return 'policy-refused'
          const unit = await obtainUnit(s.sid, sd, unitIndex, scale, 'final-raster')
          if (!alive()) return undefined
          if (typeof unit === 'string') return pageCode(unit)
          if (unit.kind !== 'raster') return 'policy-refused'
          // Fail closed: a page outside the PRINT_RASTER size contract is never marked or sent.
          if (!printSizeOk(unit, sd.doc.units[unitIndex]!, scale)) return 'policy-refused'
          const d = await deliverRaster(alive, s.sid, sd, unitIndex, scale, unit, (u) => { const f = frame(KIND.PRINT_RASTER, reqId, encodePrintBody(unitIndex, u.w, u.h, u.png)); return f ? [f] : undefined })
          return d === undefined ? undefined : typeof d === 'string' ? pageCode(d) : d.messages
        },
        send: (m) => sendFrames(ws, m),
        done: () => send(ws, { reqId, type: 'print-done', pages: n }),
        refuse: (code) => refuse(ws, reqId, code),
      })
    } finally { if (sd.print === job) sd.print = undefined }
  }

  /** A page failure: `rate-limited` (admission or lease full), anything else `policy-refused`. */
  const pageCode = (c: RefusalCode) => c === 'rate-limited' ? c : 'policy-refused'

  /** Stops the sid's pending (still validating) or running job when `reqId` names it; anything else is ignored (no reply, ever). */
  function onPrintCancel(ws: WS, msg: Record<string, unknown>) {
    const sid = ws.data.sid, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !sid || !sd || !owns(ws, sid)) return
    if (sd.print && isU32(msg.reqId) && sd.print.reqId === msg.reqId) { sd.print.stopped = true; sd.print.wake?.() }
  }

  /** The canonical-tier cache key (§9.3): never carries per-sid pixels. */
  function keyFields(doc: DocState, mark: MarkSpec, unitIndex: number, scale: number, projectionMode: 'final-raster' | 'commands', missing?: string[]): RenderKeyFields {
    return {
      docRef: doc.docRef, contentPolicyId: doc.contentPolicyId, layoutFingerprint: doc.fingerprint, recorderVersion: 1, scaleBucket: scale,
      markSettings: { visible: mark.kind, text: mark.kind === 'viewer' ? mark.text : mark.kind === 'fixed' ? 'fixed' : '', forensic: false },
      projectionMode, policyDetails: { projection: doc.contentPolicyId, ...(projectionMode === 'commands' && missing ? { missingFaces: missing } : {}) }, unitId: unitIndex,
    }
  }

  /** §9.4 steps 4–5 for one session: Layer-1 over the canonical raster under
   * the scratch lease. Unmarkable units ship canonical bytes with a reason. */
  async function markUnit(sid: string, sd: { doc: DocState; forensic?: Issued }, unitIndex: number, scale: number, unit: RenderedUnit & { kind: 'raster' }): Promise<{ unit: RenderedUnit & { kind: 'raster' }; unmarked?: string } | RefusalCode> {
    let release: () => void
    try { release = await leaseFor(sid).acquire() } catch (e) { return e instanceof LeaseFull ? 'rate-limited' : 'policy-refused' }
    try {
      const kBasis = deriveBasisKey(o.keys.kBasisRoot, sd.doc.docRef, `${unitIndex}@${scale}`)
      const r = await o.pool.encode(sid, { png: unit.png, w: unit.w, h: unit.h, kBasis, code: sd.forensic!.code, ...(o.encodeHangMs ? { hangMs: o.encodeHangMs } : {}) }, encodeEstimate(unit.w, unit.h, unit.png.length))
      if ('unmarked' in r) { forensicStats.unmarked++; return { unit, unmarked: r.unmarked } }
      forensicStats.marked++
      forensicStats.carriers.push(r.carriers); if (forensicStats.carriers.length > 256) forensicStats.carriers.shift()
      return { unit: { ...unit, png: r.png } }
    } catch (e) { return e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused' } finally { release() }
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
    if (m.type === 'search') return onSearch(ws, m)
    if (m.type === 'select') return onSelect(ws, m)
    if (m.type === 'copy') return onCopy(ws, m)
    if (m.type === 'a11y') return onA11y(ws, m)
    if (m.type === 'print') return onPrint(ws, m)
    if (m.type === 'print-cancel') return onPrintCancel(ws, m)
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
      const data: Conn = { ip, origin, url: req.url, headers: req.headers, phase: 'hello', rates: new SessionRates(), timers: new Set(), missingFaces: [], drainWaiters: new Set() }
      if (srv.upgrade(req, { data })) return undefined
      ipRates.release(ip)
      return new Response('Upgrade failed', { status: 400 })
    },
    websocket: {
      maxPayloadLength: FRAGMENT,
      drain(ws: WS) { wakeDrain(ws) },
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
        wakeDrain(ws)
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
    /** The sid's Layer-1 issuance (tests and the validation script). */
    forensicOf: (sid: string) => sessionDoc.get(sid)?.forensic,
    /** TEST-ONLY view: the canonical cached PNG this session's raster unit is built from. */
    canonicalOf: (sid: string, unitIndex: number, scale: number) => {
      const sd = sessionDoc.get(sid)
      if (!sd) return undefined
      const hit = renderCache.get(renderKey(keyFields(sd.doc, markOf(sd.policy, sd.decision), unitIndex, scale, 'final-raster')))
      return hit?.kind === 'raster' ? hit.png : undefined
    },
    table,
    stats: () => ({ sockets: sockets.size, sessions: table.size, docs: docs.size, cache: { entries: renderCache.count, bytes: renderCache.size }, pool: o.pool.stats(), forensic: { ...forensicStats, carriers: [...forensicStats.carriers] } }),
  }
  return server
}
