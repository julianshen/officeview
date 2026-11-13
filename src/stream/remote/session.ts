/**
 * The client side of the §6.2 state machine: HELLO → MANIFEST → READY →
 * STREAM (RENEW)* → CLOSE, with resume after an abnormal drop (§10).
 *
 * Engine- and React-free; the socket comes from a `StreamClient` factory so
 * tests (and hosts) can supply their own. Guards (§7.1/§14): every received
 * message is size-checked against the 1 MiB fragment cap BEFORE reassembly
 * (close 4007), only frames for an issued, still-awaited request are admitted,
 * and unit requests are admitted against the 8-outstanding / 96 MiB caps.
 * Text requests (Stage 3a search/select/copy/a11y, `text()`) keep their own
 * pending map; their answers are checked by `text-wire.ts`. One print job
 * (Stage 3b, `print()`) runs at a time; its pages are checked by `print-wire.ts`.
 */
import { decodeAssets } from '../asset-wire'
import { FRAGMENT_CAP, Kind, Reassembler, type FrameResult } from '../frame'
import { validateManifest, type Manifest } from '../manifest'
import type { Asset } from '../ops'
import { decodeRasterBody, UNIT_RASTER_DECODED_CAP } from '../raster'
import type { StreamSource } from '../source'
import { decodeUnit } from '../unit'
import { closeName, StreamSessionError } from './errors'
import { PRINT_JOB_MAX_BLOB_BYTES, PRINT_REFUSALS, printPage, printStart, type PrintedPage, type PrintRefusalCode } from './print-wire'
import { Admission, ASSET_ALLOWANCE, unitCost } from './reservation'
import { admit, emptyTally, isTerminal, requestBody, TEXT_ANSWER_TYPES, TEXT_MESSAGE_BYTES, TEXT_REFUSALS, textMessage, utf8Bytes, type TextMessage, type TextRefusalCode, type TextRequest, type TextTally } from './text-wire'

export type { RemoteRange, TextMessage, TextRequest } from './text-wire'
export type { PrintedPage } from './print-wire'

export type StreamClient = Pick<WebSocket, 'send' | 'close' | 'readyState' | 'addEventListener' | 'removeEventListener'> & { binaryType: BinaryType }
export type StreamClientFactory = (source: StreamSource) => StreamClient
export type UnitPayload =
  | { kind: 'raster'; w: number; h: number; png: Uint8Array }
  | { kind: 'commands'; ops: Uint8Array; assets: ReadonlyMap<number, Asset> }
/** A unit payload plus its admission lease: the consumer calls `release()`
 * once it has decoded/replayed the payload (idempotent). Until then the
 * request's §14 reservation stays charged. */
export type UnitLease = UnitPayload & { release(): void }
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
  /** Tightens a print job's held-blob ceiling (tests); it never rises past PRINT_JOB_MAX_BLOB_BYTES. */
  printMaxBlobBytes?: number
}
export interface PrintOptions { signal?: AbortSignal; onProgress?(done: number, total: number): void }

/** READY must reach the server inside its 2 s window; hashing gets 1.5 s of it. */
export const READY_FACE_BUDGET_MS = 1500
const RESUME_BACKOFF_MS = [250, 500, 1000]
/** PRINT_RASTER encoded cap (§14; twin of the server's PRINT_ENCODED_CAP). */
const PRINT_ENCODED_CAP = 8 << 20
const CAPS: Partial<Record<Kind, [encoded: number, decoded: number]>> = {
  [Kind.MANIFEST]: [256 << 10, 256 << 10], [Kind.ASSET_DEFINE]: [1 << 20, 8 << 20],
  [Kind.UNIT_RASTER]: [8 << 20, 64 << 20], [Kind.UNIT_COMMANDS]: [8 << 20, 8 << 20],
  [Kind.PRINT_RASTER]: [PRINT_ENCODED_CAP, 64 << 20],
}
/** The decoded bytes one page message of `from..to` may inflate to: the u32
 * unit index and 12-byte raster header plus the PNG, bounded as a unit raster
 * is (`unitCost` of the range's largest unit at scale 2), never over 64 MiB. */
function printPageCap(units: ReadonlyArray<{ w: number; h: number }>, from: number, to: number): number {
  let worst = 0
  for (let i = from; i <= to; i++) worst = Math.max(worst, unitCost(units[i]!.w, units[i]!.h, 2, false))
  return Math.min(UNIT_RASTER_DECODED_CAP, 16 + worst)
}
const UNIT_KINDS = new Set<Kind>([Kind.ASSET_DEFINE, Kind.UNIT_RASTER, Kind.UNIT_COMMANDS])
const defaultClient: StreamClientFactory = (s) => new WebSocket(s.url)

interface UnitReq {
  index: number; bucket: number; reqId: number
  /** Bytes this request was admitted for: the decoded cap of every response. */
  grant: number
  assets: Map<number, Asset>
  assetBytes: number
  resolve: (u: UnitPayload) => void; reject: (e: unknown) => void
  retried: boolean; rotatedRetry: boolean
}

interface TextReq {
  request: TextRequest; reqId: number
  on: (m: TextMessage) => void
  resolve: () => void; reject: (e: unknown) => void
  /** Sent on the current socket (false while it waits out a resume). */
  sent: boolean
  /** The next `seq` expected (copy/a11y answers). */
  seq: number
  /** What the current issue has received (checked against TEXT_LIMITS). */
  tally: TextTally
  rotatedRetry: boolean
}
interface PrintJob {
  from: number; n: number; reqId: number
  /** The decoded cap of one page message (u32 index + raster body), and the
   * admission grant the job holds for its life (one page's encoded + decoded bytes). */
  pageCap: number; grant: number
  /** Holds its admission grant (it is sent only once it does). */
  granted: boolean
  /** Sent on the current socket (false while it waits out a resume). */
  sent: boolean
  /** The checked `print-start` scale, once it has arrived. */
  scale?: 1 | 2
  /** Checked pages so far (PNG blobs only), and their bytes. */
  pages: PrintedPage[]; bytes: number
  rotatedRetry: boolean
  onProgress?: (done: number, total: number) => void
  resolve: () => void
  /** Settles once; `cancel`: a client-side failure, so the server is told to stop. */
  fail: (e: unknown, cancel: boolean) => void
}

/** Re-issued after a resume; copy and a11y are not (the server may already
 * have counted them against the copy quota). */
const REISSUED = new Set<TextRequest['type']>(['search', 'select'])

const abortError = () => Object.assign(new Error('The stream session was closed.'), { name: 'AbortError' })
const isCount = (v: unknown): v is number => Number.isInteger(v) && (v as number) >= 0
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
  /** Settles once when the session ends: the terminal error, or undefined after close(). */
  readonly ended: Promise<StreamSessionError | undefined>
  private settleEnded!: (e: StreamSessionError | undefined) => void
  private sock?: StreamClient
  private reasm!: Reassembler
  private gen = 0
  private nextReqId = 1
  private helloReqId = 0
  private readonly units = new Map<number, UnitReq>()
  private readonly texts = new Map<number, TextReq>()
  private printJob?: PrintJob
  private readonly admission = new Admission()
  private missingFaces: string[] = []
  private state: SessionState | 'opening' = 'opening'
  private failure?: StreamSessionError
  private sweeper?: ReturnType<typeof setInterval>
  private handshake?: { resolve: (m: Manifest) => void; reject: (e: unknown) => void }
  private resumeAttempt = 0
  /** Received messages are processed strictly in order (an ASSET_DEFINE's
   * hashing must finish before the UNIT_COMMANDS that references it). */
  private inbox: Promise<void> = Promise.resolve()
  private readonly manifestListeners = new Set<() => void>()

  private constructor(private readonly source: StreamSource, private readonly o: SessionOptions) {
    this.ended = new Promise((r) => { this.settleEnded = r })
  }

  /** Calls `cb` after a resume has replaced `manifest`; returns the unsubscribe. */
  onManifestChange(cb: () => void): () => void {
    this.manifestListeners.add(cb)
    return () => { this.manifestListeners.delete(cb) }
  }

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

  stats() {
    return {
      inFlight: this.admission.inFlight, inFlightBytes: this.admission.inFlightBytes, pending: this.units.size,
      printing: !!this.printJob, printHeld: this.printJob?.pages.length ?? 0, printMaxBlobBytes: this.printMaxBlobBytes(),
      /** Encoded bytes held in partial reassembly. */
      reassemblyBytes: this.reasm.bufferedBytes,
    }
  }

  /** Request one unit at a scale bucket. Settles once: payload, typed error, or abort. */
  async unit(index: number, bucket: number, signal?: AbortSignal): Promise<UnitLease> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    const dims = this.manifest.units.list[index]
    if (!dims || !this.manifest.recorder.scalesBuckets.includes(bucket)) throw new StreamSessionError('policy-refused', `no unit ${index} at bucket ${bucket}`)
    const grant = unitCost(dims.w, dims.h, bucket, this.manifest.recorder.projection !== 'final-raster')
    const release = await this.admission.acquire(grant, signal)
    if (this.isClosed()) { release(); throw this.failure ?? new StreamSessionError('closed') }
    return new Promise<UnitLease>((resolve, reject) => {
      const req: UnitReq = { index, bucket, reqId: 0, grant, assets: new Map(), assetBytes: 0, resolve: () => {}, reject, retried: false, rotatedRetry: false }
      // A delivered payload keeps the lease (its consumer releases it after
      // decode/replay); a failed or aborted request releases it here, after
      // any message for it that is still being processed (the inbox is serial).
      req.resolve = (u) => { this.units.delete(req.reqId); resolve(Object.assign(u, { release })) }
      req.reject = (e) => { this.units.delete(req.reqId); void this.inbox.then(release); reject(e) }
      signal?.addEventListener('abort', () => { if (this.units.get(req.reqId) === req) req.reject(signal.reason ?? abortError()) }, { once: true })
      if (signal?.aborted) return req.reject(signal.reason ?? abortError())
      this.issue(req)
    })
  }

  /**
   * One text request (§4.1). `on` receives each checked answer message in order;
   * the promise resolves after the terminal one (`search-done`, or the message
   * with `last: true`) and rejects with a StreamSessionError on a refusal (with
   * `limit`/`requested` when stated), a malformed answer or session end, and
   * with the abort reason on abort, after which `on` receives nothing more.
   *
   * One select is answered at a time: issuing one rejects any pending older
   * select with code 'closed' (the server sends it nothing more). On resume a
   * sent search/select is re-issued — `on` first receives `{ type: 'restart' }`
   * — while a sent copy/a11y rejects with 'session-rotated' and is never re-issued.
   */
  async text(request: TextRequest, on: (m: TextMessage) => void, signal?: AbortSignal): Promise<void> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    if (signal?.aborted) throw signal.reason ?? abortError()
    return new Promise<void>((resolve, reject) => {
      // The one settlement path: idempotent, drops the request and its abort listener.
      let settled = false
      const settle = (done: () => void) => {
        if (settled) return
        settled = true
        if (this.texts.get(req.reqId) === req) this.texts.delete(req.reqId)
        signal?.removeEventListener('abort', onAbort)
        done()
      }
      const onAbort = () => req.reject(signal!.reason ?? abortError())
      const req: TextReq = {
        request, reqId: 0, on, sent: false, seq: 0, tally: emptyTally(), rotatedRetry: false,
        resolve: () => settle(resolve),
        reject: (e) => settle(() => reject(e)),
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      if (request.type === 'select') {
        for (const r of [...this.texts.values()]) if (r.request.type === 'select') r.reject(new StreamSessionError('closed', 'superseded by a newer select'))
      }
      this.issueText(req)
    })
  }

  /**
   * One print job over units `from..to` (0-based, inclusive; §4.2). Resolves
   * with every page in unit order once `print-done` closes a job whose pages
   * all matched it; `onProgress(done, total)` follows each accepted page.
   * Anything that does not match the job exactly — a page for another unit,
   * a duplicate, out of order, outside the size contract, too many, early or
   * wrong `print-done` — fails it `policy-refused`, and nothing is kept. Over
   * PRINT_JOB_MAX_BLOB_BYTES of PNG it fails `print-too-large` with no `limit`.
   * A refusal rejects with its code (`limit`/`requested` when stated); abort
   * rejects with the abort reason; a drop mid-job rejects `session-rotated`
   * (the server stopped it; it is never re-issued); a session end rejects with
   * the session failure. Every client-side failure sends `print-cancel`. One
   * job runs at a time: another call rejects `rate-limited` without sending.
   */
  async print(from: number, to: number, opts: PrintOptions = {}): Promise<PrintedPage[]> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    const signal = opts.signal
    if (signal?.aborted) throw signal.reason ?? abortError()
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || from > to || to >= this.manifest.units.count) throw new StreamSessionError('policy-refused', `no print range ${from}..${to}`)
    if (!this.manifest.can.print) throw new StreamSessionError('policy-refused', 'print is off for this session')
    if (this.printJob) throw new StreamSessionError('rate-limited', 'a print job is already running')
    const pageCap = printPageCap(this.manifest.units.list, from, to)
    return new Promise<PrintedPage[]>((resolve, reject) => {
      // The one settlement path: idempotent, frees the slot, the held pages, the
      // abort listener, the job's partial reassembly and (once) its admission grant.
      let settled = false, release: (() => void) | undefined
      const queued = new AbortController()
      const settle = (done: (pages: PrintedPage[]) => void) => {
        if (settled) return
        settled = true
        if (this.printJob === job) this.printJob = undefined
        signal?.removeEventListener('abort', onAbort)
        queued.abort()
        if (job.reqId) this.reasm.drop(job.reqId)
        // After any message for it that is still being processed (the inbox is serial).
        if (release) { const r = release; release = undefined; void this.inbox.then(r) }
        const pages = job.pages
        job.pages = []; job.bytes = 0
        done(pages)
      }
      const onAbort = () => job.fail(signal!.reason ?? abortError(), true)
      const job: PrintJob = {
        from, n: to - from + 1, reqId: 0, pageCap, grant: PRINT_ENCODED_CAP + pageCap, granted: false,
        sent: false, pages: [], bytes: 0, rotatedRetry: false,
        ...(opts.onProgress ? { onProgress: opts.onProgress } : {}),
        resolve: () => settle(resolve),
        fail: (e, cancel) => settle(() => {
          // Best effort: the rejection runs whatever the transport does.
          try { if (cancel && job.sent && this.sock?.readyState === 1) this.send({ reqId: job.reqId, type: 'print-cancel' }) } catch { /* the socket is failing */ } finally { reject(e) }
        }),
      }
      // The slot is claimed at once (one job at a time), before admission.
      this.printJob = job
      signal?.addEventListener('abort', onAbort, { once: true })
      // §14: the job's receive/inflate workspace comes out of the session's aggregate, queued behind units.
      this.admission.acquire(job.grant, queued.signal).then((r) => {
        if (settled) return r()
        release = r; job.granted = true
        this.issuePrint(job)
      }, (e) => job.fail(e, false))
    })
  }

  /** Idempotent: closes the socket once (1000), rejects everything pending. */
  close(): void { this.end(undefined, 1000) }

  // ── internals ──────────────────────────────────────────────────────────

  private isClosed() { return this.state === 'closed' }
  private setState(s: SessionState, e?: StreamSessionError) { if (this.state === s) return; this.state = s; this.o.onState?.(s, e) }

  private issue(req: UnitReq) {
    this.units.delete(req.reqId)
    req.reqId = this.allocId(); req.assets = new Map(); req.assetBytes = 0
    this.units.set(req.reqId, req)
    // While resuming, the request waits; resume() re-issues it on the new socket.
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    this.send({ reqId: req.reqId, type: 'unit', token: this.token, unitIndex: req.index, scale: req.bucket })
    this.ensureSweeper()
  }
  private issueText(req: TextReq) {
    this.texts.delete(req.reqId)
    req.reqId = this.allocId(); req.seq = 0; req.sent = false; req.tally = emptyTally()
    this.texts.set(req.reqId, req)
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    req.sent = true
    this.send({ ...requestBody(req.request), reqId: req.reqId, type: req.request.type, token: this.token })
  }
  /** Sends the job under a fresh reqId (the retired one's partial reassembly is
   * dropped). A send that throws fails the job `policy-refused`. */
  private issuePrint(job: PrintJob) {
    if (job.reqId) this.reasm.drop(job.reqId)
    job.reqId = this.allocId(); job.sent = false
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    try {
      this.send({ reqId: job.reqId, type: 'print', token: this.token, from: job.from, to: job.from + job.n - 1 })
    } catch { return job.fail(new StreamSessionError('policy-refused', 'the print request could not be sent'), false) }
    job.sent = true
    this.ensureSweeper()
  }
  private printMaxBlobBytes() {
    const v = this.o.printMaxBlobBytes
    return typeof v === 'number' && v >= 0 && v < PRINT_JOB_MAX_BLOB_BYTES ? v : PRINT_JOB_MAX_BLOB_BYTES
  }
  /** The text messages of the running job: a refusal, `print-start`, `print-done`. */
  private onPrintText(job: PrintJob, m: Record<string, unknown>, data: string) {
    if (m.type === 'refused') {
      // A RENEW race (refused before the job was checked or charged): re-issued once.
      if (m.code === 'session-rotated' && !job.scale && !job.rotatedRetry) { job.rotatedRetry = true; return this.issuePrint(job) }
      const code: PrintRefusalCode = PRINT_REFUSALS.has(m.code as PrintRefusalCode) ? m.code as PrintRefusalCode : 'policy-refused'
      const limit = isCount(m.limit) ? m.limit : undefined, requested = isCount(m.requested) ? m.requested : undefined
      return job.fail(new StreamSessionError(code, code, { ...(limit !== undefined ? { limit } : {}), ...(requested !== undefined ? { requested } : {}) }), false)
    }
    const hostile = (why: string) => job.fail(new StreamSessionError('policy-refused', why), true)
    if (utf8Bytes(data) > TEXT_MESSAGE_BYTES) return hostile(`${String(m.type)} message over 48 KiB of UTF-8`)
    if (m.type === 'print-start') {
      if (job.scale) return hostile('a second print-start')
      const scale = printStart(m, job.n)
      if (!scale) return hostile('print-start does not match the job')
      job.scale = scale
      return
    }
    if (m.type === 'print-done') {
      if (!job.scale || job.pages.length !== job.n || m.pages !== job.n) return hostile('print-done does not close the job')
      return job.resolve()
    }
    return hostile(`unexpected ${String(m.type)} on a print job`)
  }
  /** One PRINT_RASTER of the running job: checked against the page expected next, then kept as a PNG Blob. */
  private onPrintPage(job: PrintJob, payload: Uint8Array) {
    const hostile = (why: string) => job.fail(new StreamSessionError('policy-refused', why), true)
    if (!job.scale) return hostile('a page before print-start')
    if (job.pages.length >= job.n) return hostile('more pages than announced')
    const index = job.from + job.pages.length, dims = this.manifest.units.list[index]!
    let png: Uint8Array
    try { ({ png } = printPage(payload, index, dims, job.scale)) } catch (e) { return hostile((e as Error).message) }
    job.bytes += png.byteLength
    if (job.bytes > this.printMaxBlobBytes()) return job.fail(new StreamSessionError('print-too-large', 'the print job is over the blob ceiling'), true)
    job.pages.push({ unitIndex: index, widthPx: dims.w, heightPx: dims.h, scale: job.scale, png: new Blob([png as BlobPart], { type: 'image/png' }) })
    try { job.onProgress?.(job.pages.length, job.n) } catch (e) { job.fail(e, true) }
  }

  /** A re-issued request's consumer first learns its earlier answer is void. */
  private reissueText(req: TextReq) {
    if (req.sent && !this.deliver(req, { type: 'restart' })) return
    this.issueText(req)
  }
  /** Hands `m` to the request's consumer; false once the request has settled
   * (a throwing consumer rejects it, a consumer may abort it). */
  private deliver(req: TextReq, m: TextMessage): boolean {
    try { req.on(m) } catch (e) { req.reject(e) }
    return this.texts.get(req.reqId) === req
  }
  private onTextAnswer(req: TextReq, m: Record<string, unknown>, bytes: number) {
    if (m.type === 'refused') {
      // A RENEW race (refused before the op ran): a search/select is retried once.
      if (m.code === 'session-rotated' && REISSUED.has(req.request.type) && !req.rotatedRetry) { req.rotatedRetry = true; return this.reissueText(req) }
      const code: TextRefusalCode = TEXT_REFUSALS.has(m.code as TextRefusalCode) ? m.code as TextRefusalCode : 'policy-refused'
      const limit = isCount(m.limit) ? m.limit : undefined, requested = isCount(m.requested) ? m.requested : undefined
      return req.reject(new StreamSessionError(code, code, { ...(limit !== undefined ? { limit } : {}), ...(requested !== undefined ? { requested } : {}) }))
    }
    // Over-limit answers fail closed (never truncated); the session carries on.
    if (bytes > TEXT_MESSAGE_BYTES) return req.reject(new StreamSessionError('policy-refused', `${String(m.type)} message over 48 KiB of UTF-8`))
    const msg = textMessage(req.request.type, m, req.seq, this.manifest.units.count)
    if (!msg) return req.reject(new StreamSessionError('policy-refused', `malformed ${String(m.type)} answer`))
    const over = admit(req.tally, msg)
    if (over) return req.reject(new StreamSessionError('policy-refused', over))
    if ('seq' in msg) req.seq++
    if (this.deliver(req, msg) && isTerminal(msg)) req.resolve()
  }

  private allocId() { const id = this.nextReqId; this.nextReqId = id >= 0xfffffffe ? 1 : id + 1; return id }
  private send(m: object) { this.sock?.send(JSON.stringify(m)) }

  private connect(extra: object): Promise<Manifest> {
    const gen = ++this.gen
    this.reasm = new Reassembler({
      acceptedVersions: new Set([1]),
      expects: (reqId, kind) => (kind === Kind.MANIFEST ? reqId === this.helloReqId && !!this.handshake
        // PRINT_RASTER only ever for the running, sent print job.
        : kind === Kind.PRINT_RASTER ? this.livePrint(reqId) !== undefined
        : this.units.has(reqId) && UNIT_KINDS.has(kind)),
      encodedCap: (k) => CAPS[k]?.[0] ?? 0,
      // A unit response may never decode past the bytes its request was admitted for.
      decodedCap: (k, reqId) => Math.min(CAPS[k]?.[1] ?? 0, k === Kind.MANIFEST ? Infinity
        : k === Kind.PRINT_RASTER ? this.livePrint(reqId)?.pageCap ?? 0
        : this.units.get(reqId)?.grant ?? 0),
    })
    return new Promise<Manifest>((resolve, reject) => {
      this.handshake = { resolve, reject }
      let sock: StreamClient
      try { sock = (this.o.client ?? defaultClient)(this.source) } catch {
        this.handshake = undefined
        reject(new StreamSessionError('connect-failed'))
        // A resume whose transport cannot even be built retries like a drop.
        if (this.state === 'resuming') this.scheduleResume()
        return
      }
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

  /** The running print job when `reqId` names it on the current socket. */
  private livePrint(reqId: unknown): PrintJob | undefined {
    const job = this.printJob
    return job && job.sent && job.reqId === reqId ? job : undefined
  }

  private onMessage(data: unknown) {
    // The receipt guard runs on arrival, before anything is queued or buffered.
    const size = typeof data === 'string' ? data.length : data instanceof ArrayBuffer ? data.byteLength : (data as Blob).size
    if (size > FRAGMENT_CAP) return this.end(new StreamSessionError('oversized-frame'), 4007, 'oversized-frame')
    const gen = this.gen
    this.inbox = this.inbox.then(async () => {
      if (gen !== this.gen) return
      if (typeof data === 'string') return this.onText(data)
      if (!(data instanceof ArrayBuffer)) return
      const res = await this.reasm.push(new Uint8Array(data))
      if (gen !== this.gen || !res) return
      await this.onFrame(res)
    }).catch(() => {})
  }

  private onText(data: string) {
    let m: Record<string, unknown>
    try { m = JSON.parse(data) } catch { return }
    if (m.type === 'renew' && typeof m.token === 'string' && typeof m.jti === 'string' && jtiOf(m.token) === m.jti) {
      this.send({ reqId: this.allocId(), type: 'renew-ack', jti: m.jti })
      this.token = m.token
      return
    }
    const job = this.livePrint(m.reqId)
    if (job) return this.onPrintText(job, m, data)
    const text = typeof m.reqId === 'number' ? this.texts.get(m.reqId) : undefined
    // Only text-service answers are held to 48 KiB (refused/diag are control messages).
    if (text) return this.onTextAnswer(text, m, TEXT_ANSWER_TYPES.has(m.type as string) ? utf8Bytes(data) : 0)
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
      // Fail closed: any rejected frame on the print job fails the job.
      const job = this.livePrint(res.reqId)
      if (job) return job.fail(new StreamSessionError('policy-refused', `print frame rejected: ${res.reason}`), true)
      const req = res.reqId !== undefined ? this.units.get(res.reqId) : undefined
      if (req) this.retryOrFail(req)
      else if (res.reqId !== undefined && res.reqId === this.helloReqId && this.handshake) this.end(new StreamSessionError('manifest-invalid'), 1000)
      return
    }
    if (res.kind === Kind.MANIFEST) return this.onManifest(res.payload)
    if (res.kind === Kind.PRINT_RASTER) {
      // A message that completed after its job settled is dropped.
      const job = this.livePrint(res.reqId)
      return job && this.onPrintPage(job, res.payload)
    }
    const req = this.units.get(res.reqId)
    if (!req) return
    try {
      if (res.kind === Kind.ASSET_DEFINE) {
        // §14: one request's assets stay inside the allowance it was admitted for.
        req.assetBytes += res.payload.length
        if (req.assetBytes > ASSET_ALLOWANCE) throw new Error('asset allowance exceeded')
        const defined = await decodeAssets(res.payload, req.assets)
        if (this.units.get(res.reqId) !== req) return
        for (const a of defined) req.assets.set(a.id, a)
        return
      }
      const u = decodeUnit(res.payload)
      if (u.unitId !== req.index) throw new Error('unit id mismatch')
      if (res.kind === Kind.UNIT_RASTER && u.discriminator === 2) {
        const dims = this.manifest.units.list[req.index]!
        const r = decodeRasterBody(u.body, { maxRgbaBytes: Math.min(UNIT_RASTER_DECODED_CAP, req.grant) })
        // The raster must be the unit at the requested bucket (server surfaces are ceil(w·bucket) ≤ w·bucket).
        if (r.w > dims.w * req.bucket || r.h > dims.h * req.bucket) throw new Error('raster larger than the requested unit')
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
    if (!abnormal) return this.end(new StreamSessionError('connect-failed'))
    if (hs) { this.handshake = undefined; hs.reject(new StreamSessionError('connect-failed')) }
    this.sock = undefined
    this.setState('resuming')
    // A sent copy/a11y may already have been answered and counted: never re-issued.
    for (const req of [...this.texts.values()]) {
      if (req.sent && !REISSUED.has(req.request.type)) req.reject(new StreamSessionError('session-rotated', 'the session resumed with this request in flight'))
    }
    // The server stops (and keeps charged) a job whose socket went: never re-issued.
    if (this.printJob?.sent) this.printJob.fail(new StreamSessionError('session-rotated', 'the session resumed with a print job in flight'), false)
    this.scheduleResume()
  }

  /** Next bounded resume attempt, or the terminal failure once they are spent. */
  private scheduleResume() {
    if (this.resumeAttempt >= RESUME_BACKOFF_MS.length) return this.end(new StreamSessionError('connect-failed'))
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
    // R3: a resume may change copy and text a11y; viewers re-derive what they gate on the manifest.
    for (const l of [...this.manifestListeners]) { try { l() } catch { /* a listener never breaks the resume */ } }
    for (const req of [...this.units.values()]) this.issue(req)
    // In issue order, so the newest search is the one the server keeps.
    for (const req of [...this.texts.values()]) this.reissueText(req)
    // A job called while resuming was never sent: it goes out now.
    if (this.printJob?.granted && !this.printJob.sent) this.issuePrint(this.printJob)
  }

  private ensureSweeper() {
    if (this.sweeper) return
    this.sweeper = setInterval(() => {
      for (const r of this.reasm.sweep()) {
        const job = this.livePrint(r.reqId)
        if (job) { job.fail(new StreamSessionError('policy-refused', 'a print page timed out in reassembly'), true); continue }
        const req = r.reqId !== undefined ? this.units.get(r.reqId) : undefined
        if (req) this.retryOrFail(req)
      }
      if (!this.units.size && !this.printJob && this.sweeper) { clearInterval(this.sweeper); this.sweeper = undefined }
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
    for (const req of [...this.texts.values()]) req.reject(err)
    this.printJob?.fail(err, false)
    this.admission.drain(err)
    this.setState('closed', error)
    this.settleEnded(error)
  }
}
