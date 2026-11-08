/**
 * The client side of the §6.2 state machine: HELLO → MANIFEST → READY →
 * STREAM (RENEW)* → CLOSE, with resume after an abnormal drop (§10).
 *
 * Engine- and React-free; the socket comes from a `StreamClient` factory so
 * tests (and hosts) can supply their own. Guards (§7.1/§14): every received
 * message is size-checked against the 1 MiB fragment cap BEFORE reassembly
 * (close 4007), only frames for an issued, still-awaited request are admitted,
 * and unit requests are admitted against the 8-outstanding / 96 MiB caps.
 */
import { decodeAssets } from '../asset-wire'
import { FRAGMENT_CAP, Kind, Reassembler, type FrameResult } from '../frame'
import { validateManifest, type Manifest } from '../manifest'
import type { Asset } from '../ops'
import { decodeRasterBody, UNIT_RASTER_DECODED_CAP } from '../raster'
import type { StreamSource } from '../source'
import { decodeUnit } from '../unit'
import { closeName, StreamSessionError } from './errors'
import { Admission, unitCost } from './reservation'

export type StreamClient = Pick<WebSocket, 'send' | 'close' | 'readyState' | 'addEventListener' | 'removeEventListener'> & { binaryType: BinaryType }
export type StreamClientFactory = (source: StreamSource) => StreamClient
export type UnitPayload =
  | { kind: 'raster'; w: number; h: number; png: Uint8Array }
  | { kind: 'commands'; ops: Uint8Array; assets: ReadonlyMap<number, Asset> }
/** Manifest notices, plus 'face-missing', 'text-unavailable', 'resumed', 'unit-failed'. */
export type StreamNotice = { code: string; unit?: number }
export type SessionState = 'open' | 'resuming' | 'closed'
export interface SessionOptions {
  client?: StreamClientFactory
  /** Resolves to the missing face ids; unresolved at the READY deadline ⇒ all missing. */
  faces?: (m: Manifest) => Promise<string[]>
  signal?: AbortSignal
  onNotice?: (n: StreamNotice) => void
  onState?: (s: SessionState, error?: StreamSessionError) => void
}

/** READY must reach the server inside its 2 s window; hashing gets 1.5 s of it. */
export const READY_FACE_BUDGET_MS = 1500
const RESUME_BACKOFF_MS = [250, 500, 1000]
const CAPS: Partial<Record<Kind, [encoded: number, decoded: number]>> = {
  [Kind.MANIFEST]: [256 << 10, 256 << 10], [Kind.ASSET_DEFINE]: [1 << 20, 8 << 20],
  [Kind.UNIT_RASTER]: [8 << 20, 64 << 20], [Kind.UNIT_COMMANDS]: [8 << 20, 8 << 20],
}
const UNIT_KINDS = new Set<Kind>([Kind.ASSET_DEFINE, Kind.UNIT_RASTER, Kind.UNIT_COMMANDS])
const defaultClient: StreamClientFactory = (s) => new WebSocket(s.url)

interface UnitReq {
  index: number; bucket: number; reqId: number
  assets: Map<number, Asset>
  resolve: (u: UnitPayload) => void; reject: (e: unknown) => void
  release: () => void
  retried: boolean; rotatedRetry: boolean
}

const abortError = () => Object.assign(new Error('The stream session was closed.'), { name: 'AbortError' })
export function jtiOf(token: string): string | undefined {
  try {
    const part = token.split('.')[1]!
    const json = JSON.parse(atob(part.replace(/-/g, '+').replace(/_/g, '/')))
    return typeof json.jti === 'string' ? json.jti : undefined
  } catch { return undefined }
}

export class StreamSession {
  manifest!: Manifest
  /** The current session token (rotated by RENEW, reissued by resume). */
  token = ''
  private sock?: StreamClient
  private reasm!: Reassembler
  private gen = 0
  private nextReqId = 1
  private helloReqId = 0
  private readonly units = new Map<number, UnitReq>()
  private readonly admission = new Admission()
  private missingFaces: string[] = []
  private state: SessionState | 'opening' = 'opening'
  private failure?: StreamSessionError
  private sweeper?: ReturnType<typeof setInterval>
  private handshake?: { resolve: (m: Manifest) => void; reject: (e: unknown) => void }
  private resumeAttempt = 0

  private constructor(private readonly source: StreamSource, private readonly o: SessionOptions) {}

  static async open(source: StreamSource, o: SessionOptions = {}): Promise<StreamSession> {
    const s = new StreamSession(source, o)
    if (o.signal?.aborted) throw abortError()
    o.signal?.addEventListener('abort', () => s.close(), { once: true })
    const manifest = await s.connect(source.token !== undefined ? { token: source.token } : {})
    s.manifest = manifest
    s.token = manifest.session.token
    for (const n of manifest.notices) o.onNotice?.({ code: n })
    s.missingFaces = await s.resolveFaces(manifest)
    if (s.state === 'closed') throw abortError()
    s.ready()
    s.setState('open')
    return s
  }

  stats() { return { inFlight: this.admission.inFlight, inFlightBytes: this.admission.inFlightBytes, pending: this.units.size } }

  /** Request one unit at a scale bucket. Settles once: payload, typed error, or abort. */
  async unit(index: number, bucket: number, signal?: AbortSignal): Promise<UnitPayload> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    const dims = this.manifest.units.list[index]
    if (!dims || !this.manifest.recorder.scalesBuckets.includes(bucket)) throw new StreamSessionError('policy-refused', `no unit ${index} at bucket ${bucket}`)
    const release = await this.admission.acquire(unitCost(dims.w, dims.h, bucket, this.manifest.recorder.projection !== 'final-raster'), signal)
    if (this.isClosed()) { release(); throw this.failure ?? new StreamSessionError('closed') }
    return new Promise<UnitPayload>((resolve, reject) => {
      const req: UnitReq = { index, bucket, reqId: 0, assets: new Map(), resolve, reject, release, retried: false, rotatedRetry: false }
      const settle = () => { this.units.delete(req.reqId); release() }
      req.resolve = (u) => { settle(); resolve(u) }
      req.reject = (e) => { settle(); reject(e) }
      signal?.addEventListener('abort', () => { if (this.units.get(req.reqId) === req) req.reject(signal.reason ?? abortError()) }, { once: true })
      if (signal?.aborted) return req.reject(signal.reason ?? abortError())
      this.issue(req)
    })
  }

  /** Idempotent: closes the socket once (1000), rejects everything pending. */
  close(): void { this.end(undefined, 1000) }

  // ── internals ──────────────────────────────────────────────────────────

  private isClosed() { return this.state === 'closed' }
  private setState(s: SessionState, e?: StreamSessionError) { if (this.state === s) return; this.state = s; this.o.onState?.(s, e) }

  private issue(req: UnitReq) {
    this.units.delete(req.reqId)
    req.reqId = this.allocId(); req.assets = new Map()
    this.units.set(req.reqId, req)
    // While resuming, the request waits; resume() re-issues it on the new socket.
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    this.send({ reqId: req.reqId, type: 'unit', token: this.token, unitIndex: req.index, scale: req.bucket })
    this.ensureSweeper()
  }
  private allocId() { const id = this.nextReqId; this.nextReqId = id >= 0xfffffffe ? 1 : id + 1; return id }
  private send(m: object) { this.sock?.send(JSON.stringify(m)) }

  private connect(extra: object): Promise<Manifest> {
    const gen = ++this.gen
    this.reasm = new Reassembler({
      acceptedVersions: new Set([1]),
      expects: (reqId, kind) => (kind === Kind.MANIFEST ? reqId === this.helloReqId && !!this.handshake : this.units.has(reqId) && UNIT_KINDS.has(kind)),
      encodedCap: (k) => CAPS[k]?.[0] ?? 0,
      decodedCap: (k) => CAPS[k]?.[1] ?? 0,
    })
    return new Promise<Manifest>((resolve, reject) => {
      this.handshake = { resolve, reject }
      let sock: StreamClient
      try { sock = (this.o.client ?? defaultClient)(this.source) } catch { this.handshake = undefined; return reject(new StreamSessionError('connect-failed')) }
      this.sock = sock
      sock.binaryType = 'arraybuffer'
      const live = () => gen === this.gen
      sock.addEventListener('open', () => {
        if (!live()) return
        this.helloReqId = this.allocId()
        this.send({ reqId: this.helloReqId, type: 'hello', client: 'officeview/1.0', recVersions: [1], ...extra })
      })
      sock.addEventListener('message', (e) => { if (live()) void this.onMessage((e as MessageEvent).data) })
      sock.addEventListener('close', (e) => { if (live()) this.onClose((e as CloseEvent).code) })
    })
  }

  private async resolveFaces(m: Manifest): Promise<string[]> {
    const all = m.recorder.requiredFaces.map((f) => f.id)
    if (!all.length) return []
    if (!this.o.faces) return all
    let timer: ReturnType<typeof setTimeout> | undefined
    const deadline = new Promise<string[]>((r) => { timer = setTimeout(() => r(all), READY_FACE_BUDGET_MS) })
    try {
      const missing = await Promise.race([this.o.faces(m).catch(() => all), deadline])
      return all.filter((id) => missing.includes(id))
    } finally { clearTimeout(timer) }
  }

  private ready() {
    this.send({ reqId: this.allocId(), type: 'ready', tokenJti: jtiOf(this.token) ?? '', ...(this.missingFaces.length ? { missingFaces: this.missingFaces } : {}) })
    if (this.missingFaces.length) this.o.onNotice?.({ code: 'face-missing' })
  }

  private async onMessage(data: unknown) {
    const size = typeof data === 'string' ? data.length : data instanceof ArrayBuffer ? data.byteLength : (data as Blob).size
    if (size > FRAGMENT_CAP) return this.end(new StreamSessionError('oversized-frame'), 4007, 'oversized-frame')
    if (typeof data === 'string') return this.onText(data)
    if (!(data instanceof ArrayBuffer)) return
    const gen = this.gen
    const res = await this.reasm.push(new Uint8Array(data))
    if (gen !== this.gen || !res) return
    await this.onFrame(res)
  }

  private onText(data: string) {
    let m: Record<string, unknown>
    try { m = JSON.parse(data) } catch { return }
    if (m.type === 'renew' && typeof m.token === 'string' && typeof m.jti === 'string' && jtiOf(m.token) === m.jti) {
      this.send({ reqId: this.allocId(), type: 'renew-ack', jti: m.jti })
      this.token = m.token
      return
    }
    if (m.type === 'refused' && typeof m.reqId === 'number') {
      const req = this.units.get(m.reqId)
      if (!req) return
      if (m.code === 'session-rotated' && !req.rotatedRetry) { req.rotatedRetry = true; return this.issue(req) }
      const code = m.code === 'rate-limited' || m.code === 'session-rotated' ? m.code : 'policy-refused'
      req.reject(new StreamSessionError(code))
    }
  }

  private async onFrame(res: FrameResult) {
    if (!res.ok) {
      const req = res.reqId !== undefined ? this.units.get(res.reqId) : undefined
      if (req) this.retryOrFail(req)
      else if (res.reqId !== undefined && res.reqId === this.helloReqId && this.handshake) this.end(new StreamSessionError('manifest-invalid'), 1000)
      return
    }
    if (res.kind === Kind.MANIFEST) return this.onManifest(res.payload)
    const req = this.units.get(res.reqId)
    if (!req) return
    try {
      if (res.kind === Kind.ASSET_DEFINE) {
        for (const a of await decodeAssets(res.payload, req.assets)) req.assets.set(a.id, a)
        return
      }
      const u = decodeUnit(res.payload)
      if (u.unitId !== req.index) throw new Error('unit id mismatch')
      if (res.kind === Kind.UNIT_RASTER && u.discriminator === 2) {
        const r = decodeRasterBody(u.body, { maxRgbaBytes: UNIT_RASTER_DECODED_CAP })
        return req.resolve({ kind: 'raster', w: r.w, h: r.h, png: r.png })
      }
      if (res.kind === Kind.UNIT_COMMANDS && u.discriminator === 1) return req.resolve({ kind: 'commands', ops: u.body, assets: req.assets })
      throw new Error('unexpected unit kind')
    } catch {
      if (this.units.get(res.reqId) === req) this.retryOrFail(req)
    }
  }

  /** §7.6/§10: a unit that fails validation is re-requested once, then fails. */
  private retryOrFail(req: UnitReq) {
    if (req.retried) { this.o.onNotice?.({ code: 'unit-failed', unit: req.index }); return req.reject(new StreamSessionError('unit-failed')) }
    req.retried = true
    this.issue(req)
  }

  private onManifest(payload: Uint8Array) {
    const hs = this.handshake
    if (!hs) return
    this.handshake = undefined
    let v: ReturnType<typeof validateManifest>
    try { v = validateManifest(JSON.parse(new TextDecoder().decode(payload))) } catch { v = { ok: false, field: '$' } }
    if (!v.ok) { hs.reject(new StreamSessionError('manifest-invalid', `manifest field ${v.field}`)); return this.end(new StreamSessionError('manifest-invalid'), 1000) }
    hs.resolve(v.manifest)
  }

  private onClose(code: number) {
    const hs = this.handshake
    const name = closeName(code)
    if (this.state === 'closed') return
    if (name && name !== 'oversized-frame') return this.end(new StreamSessionError(name))
    const abnormal = code === 1001 || code === 1006 || code === 1000 || code === 1005
    // Before the first MANIFEST there is nothing to resume.
    if (!this.manifest) { hs?.reject(new StreamSessionError('connect-failed')); this.handshake = undefined; return this.end(new StreamSessionError('connect-failed')) }
    if (!abnormal || this.resumeAttempt >= RESUME_BACKOFF_MS.length) return this.end(new StreamSessionError('connect-failed'))
    if (hs) { this.handshake = undefined; hs.reject(new StreamSessionError('connect-failed')) }
    this.setState('resuming')
    const delay = RESUME_BACKOFF_MS[this.resumeAttempt++]!
    setTimeout(() => { if (this.state === 'resuming') void this.resume() }, delay)
  }

  private async resume() {
    const prior = this.manifest
    let m: Manifest
    try { m = await this.connect({ resume: { sid: prior.session.sid, resumeProof: this.token } }) } catch { return }
    if (this.state !== 'resuming') return
    if (m.docRef !== prior.docRef || m.layoutFingerprint !== prior.layoutFingerprint || m.session.sid !== prior.session.sid || m.units.count !== prior.units.count) {
      return this.end(new StreamSessionError('layout-changed'), 1000)
    }
    this.manifest = m
    this.token = m.session.token
    this.resumeAttempt = 0
    this.ready()
    this.setState('open')
    this.o.onNotice?.({ code: 'resumed' })
    for (const req of [...this.units.values()]) this.issue(req)
  }

  private ensureSweeper() {
    if (this.sweeper) return
    this.sweeper = setInterval(() => {
      for (const r of this.reasm.sweep()) { const req = r.reqId !== undefined ? this.units.get(r.reqId) : undefined; if (req) this.retryOrFail(req) }
      if (!this.units.size && this.sweeper) { clearInterval(this.sweeper); this.sweeper = undefined }
    }, 1000)
  }

  private end(error: StreamSessionError | undefined, closeCode?: number, reason?: string) {
    if (this.state === 'closed') return
    this.gen++
    this.failure = error
    if (this.sweeper) { clearInterval(this.sweeper); this.sweeper = undefined }
    const sock = this.sock
    this.sock = undefined
    if (sock && closeCode !== undefined && sock.readyState < 2) sock.close(closeCode, reason)
    const err = error ?? abortError()
    if (this.handshake) { this.handshake.reject(err); this.handshake = undefined }
    for (const req of [...this.units.values()]) req.reject(err)
    this.admission.drain(err)
    this.setState('closed', error)
  }
}
