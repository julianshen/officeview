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
 * A presenter's notes image (Stage 3c, `notes()`) is one NOTES_RASTER per
 * request, admitted only for that live, sent request.
 * Unit sends (R28) are held to the session's one send log (`unitBudget`), which
 * lives as long as the server sid: a first send waits, first in first out, for
 * room under UNIT_FIRST_CAP; a re-send for room under UNIT_RESEND_CAP; nothing
 * charged is sent while a `rate-limited` refusal blocks the log.
 */
import { decodeAssets } from '../asset-wire'
import { FRAGMENT_CAP, HEADER_BYTES, Kind, Reassembler, REASSEMBLY_MS, type FrameResult } from '../frame'
import { validateManifest, type Manifest } from '../manifest'
import type { Asset } from '../ops'
import { decodeRasterBody, UNIT_RASTER_DECODED_CAP } from '../raster'
import { decodeShowMessage, SHOW_PLAN_LIMITS } from '../show-plan'
import type { AllocationOwner, Lease } from '../../slideshow/types'
import type { StreamSource } from '../source'
import { decodeUnit } from '../unit'
import { closeName, StreamSessionError, type StreamErrorCode } from './errors'
import { PRINT_JOB_MAX_BLOB_BYTES, PRINT_JOB_MAX_DECODED, PRINT_REFUSALS, printDecodedBound, printPage, printStart, type PrintedPage, type PrintRefusalCode } from './print-wire'
import { Admission, ASSET_ALLOWANCE, SESSION_TRANSIENT_BYTES, unitCost } from './reservation'
import { UNIT_FIRST_CAP, UNIT_RESEND_CAP, UNIT_SPAN_MS, UnitSendLog, type UnitBudgetView } from './unit-budget'
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
/** A checked notes image plus its admission lease (released once decoded; idempotent). */
export type NotesLease = { w: number; h: number; png: Uint8Array; release(): void }
/** A show reply's kind-6 payloads in arrival order, plus its admission lease (released once decoded; idempotent). */
export type ShowReply = { messages: Uint8Array[]; release(): void }
export type ShowMotion = 'full' | 'reduced'
/** What a show request needs besides its fields: the show workspace its retained payloads are charged to, and R16's hook. */
export interface ShowRequestOptions {
  /** Each completed message is leased here at its encoded bytes until `decodeBundle` or `release()` frees it (ruling 5). */
  workspace: AllocationOwner
  /** R16: called once per server-charged send. */
  onSend?: () => void
}
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
  /** The monotonic ms clock of reassembly and request deadlines (tests); default performance.now. */
  now?: () => number
  /** Tightens a print job's held-blob ceiling (tests); it never rises past PRINT_JOB_MAX_BLOB_BYTES. */
  printMaxBlobBytes?: number
}
export interface PrintOptions { signal?: AbortSignal; onProgress?(done: number, total: number): void }

/** READY must reach the server inside its 2 s window; hashing gets 1.5 s of it. */
export const READY_FACE_BUDGET_MS = 1500
const RESUME_BACKOFF_MS = [250, 500, 1000]
/** PRINT_RASTER encoded cap (§14; twin of the server's PRINT_ENCODED_CAP). */
const PRINT_ENCODED_CAP = 8 << 20
/** NOTES_RASTER encoded cap (§14; twin of the server's NOTES_ENCODED_CAP). */
const NOTES_ENCODED_CAP = 8 << 20
export const CAPS: Partial<Record<Kind, [encoded: number, decoded: number]>> = {
  [Kind.MANIFEST]: [256 << 10, 256 << 10], [Kind.ASSET_DEFINE]: [1 << 20, 8 << 20],
  [Kind.UNIT_RASTER]: [8 << 20, 64 << 20], [Kind.UNIT_COMMANDS]: [8 << 20, 8 << 20],
  [Kind.PRINT_RASTER]: [PRINT_ENCODED_CAP, 64 << 20],
  [Kind.NOTES_RASTER]: [NOTES_ENCODED_CAP, 64 << 20],
  // SHOW_PLAN (piece 4): one message ≤ 8 MiB encoded / 32 MiB decoded (parent §14); a bundle is several messages.
  [Kind.SHOW_PLAN]: [8 << 20, 32 << 20],
  // MEDIA_CHUNK (piece 5, D14): stored, so its decoded cap equals its encoded cap. No request expects one yet.
  [Kind.MEDIA_CHUNK]: [4 << 20, 4 << 20],
}
/** The notes image (twins of server/src/notes.ts): exactly 720 × scale device px
 * wide, 1..NOTES_MAX_PX device px tall. */
const NOTES_WIDTH_CSS = 720, NOTES_MAX_PX = 8192
/** The decoded cap of the NOTES_RASTER of a declared `w × h` image (R10): the
 * u32 unit index and the 12-byte raster header plus the image's RGBA (a PNG of
 * text on white is far smaller; one that does not compress fails closed),
 * never over 64 MiB. */
export const notesBodyCap = (w: number, h: number) => Math.min(UNIT_RASTER_DECODED_CAP, 16 + w * h * 4)
/** The admission grant of a declared `w × h` notes image (R10), its true peak
 * as print's: the encoded bound (the body plus deflate/gzip slack, as
 * `unitCost`; `notes-start` declares no encoded size), at most the 8 MiB kind
 * cap, plus 2 × the decoded body (inflated chunks + their contiguous copy).
 * Over the 96 MiB aggregate (the tallest scale-2 images, ≈ 98 MiB) Admission
 * clamps it to the aggregate. */
export const notesGrant = (w: number, h: number) => {
  const body = notesBodyCap(w, h)
  return Math.min(NOTES_ENCODED_CAP, Math.ceil(body * 1.01) + 65536) + 2 * body
}
/** The refusal codes a notes request rejects with; any other is `policy-refused`. */
const NOTES_REFUSALS = new Set<StreamErrorCode>(['policy-refused', 'invalid-unit', 'rate-limited', 'session-rotated'])
/** The refusal codes a show request rejects with (server/src/ws.ts ShowErrorCode); any other is `policy-refused`.
 * `still`, `show-budget`, `show-limits`, `show-plan-timeout`: play the piece-3 still; `morph-fallback`: the fade. */
const SHOW_REFUSALS = new Set<StreamErrorCode>([...NOTES_REFUSALS, 'still', 'show-budget', 'show-limits', 'show-plan-timeout', 'morph-fallback'])
/** A `*-start` header is one text message ≤ 64 KiB (D1). */
const SHOW_START_BYTES = 64 << 10
/** A sent show request with no start yet fails after this (`now()`): the server's build has a 10 s
 * deadline (SHOW_PLAN_DEADLINE_MS) plus queueing and marking, and refuses `show-plan-timeout` itself. */
export const SHOW_REPLY_MS = 30_000
/** The most kind-6 messages a reply may declare (D1, D2): a plan, its entry frame and one asset per
 * layer; a track's frames; a Morph record and its assets (every side and `over` asset) or frames. */
const SHOW_MAX_MESSAGES = {
  plan: 2 + SHOW_PLAN_LIMITS.layers,
  track: SHOW_PLAN_LIMITS.trackFrames,
  morph: 1 + 3 * SHOW_PLAN_LIMITS.morphItems,
} as const
/** Per encoded message, the slack a retained reply may run past its declared decoded bytes (an
 * incompressible PNG is slightly larger than its RGBA). */
const SHOW_MESSAGE_SLACK = 65536
/**
 * The admission grant of a show reply declaring `decodedBytes` (D1): one message is in reassembly at
 * a time, so the grant is one message's true peak, as print's: its encoded bound (≤ the 8 MiB kind
 * cap) plus 2 × its decoded body (inflated chunks + their contiguous copy), the body at most the
 * 32 MiB kind-6 decoded cap. Controller ruling (Task 10 follow-up): it covers only the message in
 * flight — taken at its first fragment, released when it completes reassembly.
 */
export const showGrant = (decodedBytes: number) => showEncodedAllowance(decodedBytes) + 2 * Math.min(CAPS[Kind.SHOW_PLAN]![1], decodedBytes)
/** The grant's encoded part: the most encoded bytes one message of the reply may declare (enforced before buffering). */
export const showEncodedAllowance = (decodedBytes: number) => Math.min(CAPS[Kind.SHOW_PLAN]![0], Math.ceil(decodedBytes * 1.01) + 65536)
/**
 * A print job's peak transient bytes for one page are its encoded fragments
 * (`Reassembler.push`, src/stream/frame.ts:124, ≤ PRINT_ENCODED_CAP) plus the
 * inflated chunks (frame.ts:161) plus their contiguous copy (frame.ts:164-165):
 * encoded + 2 × decoded. The job's grant is that peak, so the decoded page
 * body is clamped to keep the grant inside the session's aggregate admission.
 * A page whose body is larger (a PNG over ~44 MiB) cannot print on the stream path.
 */
export const PRINT_PAGE_BODY_MAX = Math.floor((SESSION_TRANSIENT_BYTES - PRINT_ENCODED_CAP) / 2)
/** The admission grant of a job whose page bodies may decode to `pageCap` bytes (see PRINT_PAGE_BODY_MAX). */
export const printGrant = (pageCap: number) => PRINT_ENCODED_CAP + 2 * pageCap
/** The decoded bytes one page message of `from..to` may inflate to: the u32
 * unit index and 12-byte raster header plus the PNG, bounded as a unit raster
 * is (`unitCost` of the range's largest unit at scale 2), never over 64 MiB
 * nor PRINT_PAGE_BODY_MAX. */
export function printPageCap(units: ReadonlyArray<{ w: number; h: number }>, from: number, to: number): number {
  let worst = 0
  for (let i = from; i <= to; i++) worst = Math.max(worst, unitCost(units[i]!.w, units[i]!.h, 2, false))
  return Math.min(UNIT_RASTER_DECODED_CAP, PRINT_PAGE_BODY_MAX, 16 + worst)
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
  /** R16: called once per server-charged send (not the stale-jti re-send). */
  onSend?: () => void
  /** A charged send of it has gone out (later ones are re-sends, R28). */
  sent: boolean
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
  /** Checked pages so far (PNG blobs only), their PNG bytes, and their decoded bytes (Σ w·h·4). */
  pages: PrintedPage[]; bytes: number; decoded: number
  rotatedRetry: boolean
  onProgress?: (done: number, total: number) => void
  resolve: () => void
  /** Settles once; `cancel`: a client-side failure, so the server is told to stop. */
  fail: (e: unknown, cancel: boolean) => void
}

interface NotesReq {
  unit: number; scale: 1 | 2; reqId: number
  /** The decoded cap of its one NOTES_RASTER (see notesBodyCap); 0 until `notes-start` is granted. */
  bodyCap: number
  /** Takes the admission grant of the declared image (released on settle unless leased out). */
  hold: (release: () => void) => void
  /** Sent on the current socket (false while it waits out a resume). */
  sent: boolean
  /** The checked `notes-start` size, once it has arrived. */
  start?: { w: number; h: number }
  rotatedRetry: boolean
  /** While no message for it is in reassembly, it fails at this time (`now()`):
   * REASSEMBLY_MS after the send, restarted at `notes-start`. */
  deadline: number
  resolve: (img: { w: number; h: number; png: Uint8Array }) => void
  fail: (e: unknown) => void
}

type ShowKind = 'plan' | 'track' | 'morph'
interface ShowReq {
  readonly show: ShowKind
  /** The request's own fields (`unit` [+ `step`], or `from`, `to`, `fromStep`), sent and echoed by the start. */
  readonly fields: Readonly<Record<string, number>>
  /** The server unit every asset of the reply carries (a Morph pair: `to`). */
  readonly unit: number
  readonly scale: 1 | 2; readonly motion: ShowMotion
  reqId: number
  /** Sent on the current socket (false while it waits for log room or a resume). */
  sent: boolean
  /** The checked start: declared messages and decoded bytes, one message's decoded cap, the retained-bytes bound. */
  start?: { messages: number; decodedBytes: number; encodedCap: number; bodyCap: number; retainCap: number }
  /** A Morph reply's record says it is composited: its assets are frames (tag 4). */
  composited?: boolean
  /** The checked kind-6 payloads so far, and their bytes. */
  received: Uint8Array[]; bytes: number
  /** Every declared message has arrived: the reply settles once the inbox has drained what was queued with it. */
  complete: boolean
  rotatedRetry: boolean
  /** While no message for it is in reassembly, it fails at this time (`now()`). */
  deadline: number
  /** R16: called once per server-charged send. */
  onSend?: () => void
  readonly workspace: AllocationOwner
  /** The admission grant of the message in flight (one at a time), released when it completes. */
  grant?: () => void
  /** The workspace leases of the completed messages (their encoded bytes). */
  leases: Lease[]
  /** Admission could not fit its next message: its rejection surfaces as `admission`. */
  admissionShort: boolean
  resolve: (messages: Uint8Array[]) => void
  fail: (e: unknown) => void
}
const isShow = (r: UnitReq | ShowReq): r is ShowReq => 'show' in r

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
  /** Granted notes requests by their current reqId. */
  private readonly notesReqs = new Map<number, NotesReq>()
  /** Show requests (piece 4) by their current reqId. */
  private readonly shows = new Map<number, ShowReq>()
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
  /** R28: every server-charged unit send of this sid, whoever asked (survives resume). */
  private readonly sendLog = new UnitSendLog(() => this.now(), UNIT_FIRST_CAP, UNIT_SPAN_MS)
  /** Unit and show requests waiting for log room, oldest first (each still in `units` / `shows`). */
  private readonly waiting = new Set<UnitReq | ShowReq>()
  private drainTimer?: ReturnType<typeof setTimeout>
  /**
   * R28: the session's unit send budget, for callers that pace themselves (the
   * show fetcher). Room is UNIT_FIRST_CAP less the window's charged sends and
   * the first sends already waiting here; `pending` is the caller's own share.
   */
  readonly unitBudget: UnitBudgetView = {
    room: (pending = 0) => this.sendLog.room(pending + this.waitingFirst()),
    waitMs: (pending = 0) => this.sendLog.waitMs(pending + this.waitingFirst()),
  }

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
      printing: !!this.printJob, notesPending: this.notesReqs.size, showPending: this.shows.size, printHeld: this.printJob?.pages.length ?? 0, printMaxBlobBytes: this.printMaxBlobBytes(),
      /** Encoded bytes held in partial reassembly. */
      reassemblyBytes: this.reasm.bufferedBytes,
    }
  }

  /**
   * Request one unit at a scale bucket. Settles once: payload, typed error, or abort.
   * `onSend` (R16) is called on every send the server charges to the `'unit'` rate
   * class: the first, a validation re-send and a resume re-send — never the
   * stale-jti re-send (refused before it was charged). Every such send is also
   * stamped on the session's send log (R28); over budget, the request waits for room.
   */
  async unit(index: number, bucket: number, signal?: AbortSignal, onSend?: () => void): Promise<UnitLease> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    const dims = this.manifest.units.list[index]
    if (!dims || !this.manifest.recorder.scalesBuckets.includes(bucket)) throw new StreamSessionError('policy-refused', `no unit ${index} at bucket ${bucket}`)
    const grant = unitCost(dims.w, dims.h, bucket, this.manifest.recorder.projection !== 'final-raster')
    const release = await this.admission.acquire(grant, signal)
    if (this.isClosed()) { release(); throw this.failure ?? new StreamSessionError('closed') }
    return new Promise<UnitLease>((resolve, reject) => {
      const req: UnitReq = { index, bucket, reqId: 0, grant, assets: new Map(), assetBytes: 0, resolve: () => {}, reject, retried: false, rotatedRetry: false, sent: false, ...(onSend ? { onSend } : {}) }
      // A delivered payload keeps the lease (its consumer releases it after
      // decode/replay); a failed or aborted request releases it here, after
      // any message for it that is still being processed (the inbox is serial).
      req.resolve = (u) => { this.units.delete(req.reqId); this.waiting.delete(req); resolve(Object.assign(u, { release })) }
      req.reject = (e) => { this.units.delete(req.reqId); this.waiting.delete(req); void this.inbox.then(release); reject(e) }
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
   * PRINT_JOB_MAX_BLOB_BYTES of PNG, or a `print-start` whose scale puts the
   * job's decoded images (Σ w·h·4 from the manifest sizes) over
   * PRINT_JOB_MAX_DECODED, it fails `print-too-large` with no `limit`.
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
        job.pages = []; job.bytes = 0; job.decoded = 0
        done(pages)
      }
      const onAbort = () => job.fail(signal!.reason ?? abortError(), true)
      const job: PrintJob = {
        from, n: to - from + 1, reqId: 0, pageCap, grant: printGrant(pageCap), granted: false,
        sent: false, pages: [], bytes: 0, decoded: 0, rotatedRetry: false,
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

  /**
   * A presenter's notes image of one streamed slide at scale 1 or 2 (§12a.5,
   * §12a.8): `notes-start {w, h, scale}` then one NOTES_RASTER for the unit,
   * both checked against the request (width exactly 720 × scale, height
   * 1..8192, the unit asked, the announced size); anything else rejects
   * `policy-refused`. Refuses without sending when `can.notes` is off
   * (`policy-refused`) or the unit is out of range or hidden (`invalid-unit`).
   * A refusal rejects with its code; a `session-rotated` one (stale jti) is
   * re-issued once. Abort rejects with its reason; a drop with the request
   * sent rejects `session-rotated` (never re-issued: the server audits each);
   * a session end rejects with the session failure. Nothing is reserved until
   * `notes-start` declares the image (R10); its grant (`notesGrant`) is taken
   * then or never: no room rejects `rate-limited`. The lease keeps the grant
   * until `release()`.
   */
  async notes(unit: number, scale: 1 | 2, signal?: AbortSignal): Promise<NotesLease> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    if (signal?.aborted) throw signal.reason ?? abortError()
    const m = this.manifest
    if (!m.can.notes) throw new StreamSessionError('policy-refused', 'notes are off for this session')
    if ((scale !== 1 && scale !== 2) || !m.recorder.scalesBuckets.includes(scale)) throw new StreamSessionError('policy-refused', `no notes scale ${scale}`)
    const dims = Number.isInteger(unit) && unit >= 0 ? m.units.list[unit] : undefined
    if (!dims || dims.hidden) throw new StreamSessionError('invalid-unit', `no notes for unit ${unit}`)
    return new Promise<NotesLease>((resolve, reject) => {
      // The one settlement path: idempotent, drops the request, its abort
      // listener, its partial reassembly and — unless leased out — its grant.
      let settled = false, release: (() => void) | undefined
      const settle = (done: () => void, keepGrant = false) => {
        if (settled) return
        settled = true
        if (this.notesReqs.get(req.reqId) === req) this.notesReqs.delete(req.reqId)
        signal?.removeEventListener('abort', onAbort)
        if (req.reqId) this.reasm.drop(req.reqId)
        // After any message for it that is still being processed (the inbox is serial).
        if (release && !keepGrant) { const r = release; release = undefined; void this.inbox.then(r) }
        done()
      }
      const onAbort = () => req.fail(signal!.reason ?? abortError())
      const req: NotesReq = {
        unit, scale, reqId: 0, bodyCap: 0, sent: false, rotatedRetry: false, deadline: 0,
        hold: (r) => { if (settled) r(); else release = r },
        resolve: (img) => settle(() => {
          const r = release!
          resolve({ ...img, release: () => { void this.inbox.then(r) } })
        }, true),
        fail: (e) => settle(() => reject(e)),
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.issueNotes(req)
    })
  }

  /**
   * Piece 4 (D1): one slide bundle of a streamed show. The reply is `show-plan-start {unit, messages,
   * decodedBytes}` (any `diag` before it is tolerated), then exactly `messages` SHOW_PLAN (kind 6)
   * payloads: the plan record, the entry frame, the layers. Each is checked against the request (the
   * tag order of its kind, an asset's unit) and kept as received; `decodeBundle` (show-bundle.ts)
   * turns them into bitmaps. Refuses without sending when `can.present` is off, the scale or motion
   * is not one the server takes (`policy-refused`), or the unit is out of range or hidden
   * (`invalid-unit`). A refusal rejects with its code: `still`, `show-budget`, `show-limits` and
   * `show-plan-timeout` mean "play the piece-3 still"; an unknown code is `policy-refused`; a
   * `session-rotated` one (stale jti, uncharged) is re-sent once. Anything else that does not match
   * the request — a start over 64 KiB, for another unit or kind, declaring over 128 MiB or an
   * impossible count; a message before the start, past the declared count, out of order or for
   * another unit — rejects `policy-refused`, and everything it holds is released. Admission
   * (`showGrant`) covers only the message in flight: taken at its first fragment (now or never: no
   * room fails the reply with the local, transient `admission`, never a server rate pause) and
   * released once it completes. A completed message is leased in `o.workspace` at its encoded bytes
   * (no room: `workspace-full`) until `decodeBundle` or the reply's `release()` frees it.
   * The send is charged to the `'unit'` class: it is held to the session's send log like a unit
   * request (R28), recorded once, and `onSend` (R16) is called once. Abort rejects with its reason;
   * a drop with the request sent rejects `session-rotated` (never re-issued: the server charged and
   * audited it); a session end rejects with the session failure.
   */
  showPlan(unit: number, scale: 1 | 2, motion: ShowMotion, signal: AbortSignal, o: ShowRequestOptions): Promise<ShowReply> {
    return this.showRequest('plan', unit, { unit }, scale, motion, signal, o)
  }
  /**
   * One keyframe track of `step` (D1, D4): `show-track-start {unit, step, messages, decodedBytes, encodedBytes}`, then
   * tag-4 frames only. D23 (final review finding 1): the frames are never all decoded together, so `decodedBytes` is
   * the live window's bound (at most 8 frames' `w·h·4`) and `encodedBytes` is Σ the encoded payloads, which stay
   * retained while the track plays: each ≤ 128 MiB, the encoded total also ≤ `messages` × 8 MiB, and the payloads
   * received never past it. Otherwise as `showPlan`; a track or Morph start of any other key set is refused.
   */
  showTrack(unit: number, step: number, scale: 1 | 2, motion: ShowMotion, signal: AbortSignal, o: ShowRequestOptions): Promise<ShowReply> {
    const bad = !Number.isInteger(step) || step < 0 || step >= SHOW_PLAN_LIMITS.steps
    return this.showRequest('track', unit, { unit, step }, scale, motion, signal, o, bad ? `no step ${step}` : undefined)
  }
  /**
   * One Morph pair (D1, D11): `show-morph-start`, then the tag-5 record, whose `from`, `to` and
   * `fromStep` must be the request's, then its tag-2 assets or, for a composited record, exactly its
   * `count` tag-4 frames. Every asset carries `to`'s unit. `morph-fallback` means play the fade. The start also declares
   * `encodedBytes`, as a track's; a composited pair's `decodedBytes` is its window bound plus the record, a layered
   * pair's Σ its assets plus the record (all held together). As `showPlan`.
   */
  showMorph(from: number, to: number, fromStep: number, scale: 1 | 2, motion: ShowMotion, signal: AbortSignal, o: ShowRequestOptions): Promise<ShowReply> {
    const bad = !this.shownUnit(from) || from === to || !Number.isInteger(fromStep) || fromStep < 0 || fromStep > SHOW_PLAN_LIMITS.steps
    return this.showRequest('morph', to, { from, to, fromStep }, scale, motion, signal, o, bad ? `no Morph pair ${from} → ${to} at ${fromStep}` : undefined)
  }

  private shownUnit(u: number) {
    const d = Number.isInteger(u) && u >= 0 ? this.manifest.units.list[u] : undefined
    return !!d && !d.hidden
  }
  /** The one show request path (see `showPlan`); `invalid`: why the request's own fields name nothing (`invalid-unit`). */
  private async showRequest(show: ShowKind, unit: number, fields: Record<string, number>, scale: 1 | 2, motion: ShowMotion, signal: AbortSignal, o: ShowRequestOptions, invalid?: string): Promise<ShowReply> {
    if (this.state === 'closed') throw this.failure ?? new StreamSessionError('closed')
    if (signal?.aborted) throw signal.reason ?? abortError()
    const m = this.manifest
    if (!m.can.present) throw new StreamSessionError('policy-refused', 'presenting is off for this session')
    if ((scale !== 1 && scale !== 2) || !m.recorder.scalesBuckets.includes(scale)) throw new StreamSessionError('policy-refused', `no show scale ${scale}`)
    if (motion !== 'full' && motion !== 'reduced') throw new StreamSessionError('policy-refused', `no show motion ${String(motion)}`)
    if (invalid || !this.shownUnit(unit)) throw new StreamSessionError('invalid-unit', invalid ?? `no show unit ${unit}`)
    // D13: under reduced motion a Morph plays as a fade; the server refuses the pair (`invalid-unit`), so it is never sent.
    if (show === 'morph' && motion === 'reduced') throw new StreamSessionError('invalid-unit', 'a Morph pair under reduced motion plays as a fade')
    return new Promise<ShowReply>((resolve, reject) => {
      // The one settlement path: idempotent, drops the request, its place in the send queue, its
      // abort listener, its partial reassembly, its message grant and — unless handed out — its payload leases.
      let settled = false
      const settle = (done: () => void, keepLeases = false) => {
        if (settled) return
        settled = true
        if (this.shows.get(req.reqId) === req) this.shows.delete(req.reqId)
        this.waiting.delete(req)
        signal?.removeEventListener('abort', onAbort)
        if (req.reqId) this.reasm.drop(req.reqId)
        // After any message for it that is still being processed (the inbox is serial).
        if (req.grant) { const r = req.grant; req.grant = undefined; void this.inbox.then(r) }
        if (!keepLeases) for (const l of req.leases.splice(0)) l.release()
        req.received = []; req.bytes = 0
        done()
      }
      const onAbort = () => req.fail(signal!.reason ?? abortError())
      const req: ShowReq = {
        show, fields, unit, scale, motion, reqId: 0, sent: false, received: [], bytes: 0, complete: false, rotatedRetry: false, deadline: 0,
        workspace: o.workspace, leases: [], admissionShort: false,
        ...(o.onSend ? { onSend: o.onSend } : {}),
        resolve: (messages) => settle(() => {
          const leases = req.leases.splice(0)
          let freed = false
          resolve({ messages, release: () => { if (freed) return; freed = true; for (const l of leases) l.release() } })
        }, true),
        fail: (e) => settle(() => reject(e)),
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.issueShow(req)
    })
  }

  /** Idempotent: closes the socket once (1000), rejects everything pending. */
  close(): void { this.end(undefined, 1000) }

  // ── internals ──────────────────────────────────────────────────────────

  private isClosed() { return this.state === 'closed' }
  private setState(s: SessionState, e?: StreamSessionError) { if (this.state === s) return; this.state = s; this.o.onState?.(s, e) }

  /** `charged` false only for the stale-jti re-send (R16: refused before the server charged it). */
  private issue(req: UnitReq, charged = true) {
    this.units.delete(req.reqId)
    const prev = req.reqId
    req.reqId = this.allocId(); req.assets = new Map(); req.assetBytes = 0
    // R30: the uncharged re-send takes over the refused send's slot (one charge, aged from its own reply).
    if (!charged) this.sendLog.transfer(prev, req.reqId)
    this.units.set(req.reqId, req)
    // While resuming, the request waits; resume() re-issues it on the new socket.
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    // R28: a first send queues behind any already waiting; a re-send needs only room.
    if (charged && ((!req.sent && this.waiting.size) || !this.hasRoom(req))) { this.waiting.add(req); return this.scheduleDrain() }
    this.transmit(req, charged)
  }
  private transmit(req: UnitReq, charged: boolean) {
    this.waiting.delete(req)
    this.send({ reqId: req.reqId, type: 'unit', token: this.token, unitIndex: req.index, scale: req.bucket })
    if (charged) {
      req.sent = true
      // R29: it holds its slot until the server's first reply for this reqId, then a span.
      this.sendLog.record(req.reqId)
      try { req.onSend?.() } catch { /* a hook never breaks the send */ }
    }
    this.ensureSweeper()
  }
  /** R29: the first server reply for a unit send starts its span; a waiting send may now get a timed slot. */
  private unitAnswered(reqId: number) {
    if (!this.sendLog.has(reqId)) return
    this.sendLog.answered(reqId)
    if (this.drainTimer && this.waiting.size) { clearTimeout(this.drainTimer); this.drainTimer = undefined; this.scheduleDrain() }
  }
  /** A `session-rotated` refusal of a live unit request that onText re-sends once. */
  private rotatedResend(m: Record<string, unknown>) {
    if (m.code !== 'session-rotated') return false
    // A show request's stale-token refusal comes as `error` (R30: its slot passes to the re-send).
    if (m.type === 'error') { const s = this.liveShow(m.reqId); return !!s && !s.start && !s.rotatedRetry }
    if (m.type !== 'refused') return false
    const req = this.units.get(m.reqId as number)
    return !!req && !req.rotatedRetry
  }
  private hasRoom(req: UnitReq | ShowReq) { return this.sendLog.room(0, req.sent ? UNIT_RESEND_CAP : UNIT_FIRST_CAP) > 0 }
  private waitingFirst() { let n = 0; for (const r of this.waiting) if (!r.sent) n++; return n }
  /** Sends what the log now has room for, oldest first, then waits for the next slot. */
  private drain() {
    this.drainTimer = undefined
    // While resuming, resume() re-issues every request (waiting ones included).
    if (this.state !== 'open' || !this.sock || this.sock.readyState !== 1) return
    for (const req of this.waiting) {
      if ((isShow(req) ? this.shows : this.units).get(req.reqId) !== req) { this.waiting.delete(req); continue }
      if (!this.hasRoom(req)) break
      if (isShow(req)) this.transmitShow(req, true)
      else this.transmit(req, true)
    }
    this.scheduleDrain()
  }
  private scheduleDrain() {
    const head = this.waiting.values().next().value
    if (this.drainTimer || !head) return
    const ms = this.sendLog.waitMs(0, head.sent ? UNIT_RESEND_CAP : UNIT_FIRST_CAP)
    this.drainTimer = setTimeout(() => this.drain(), Number.isFinite(ms) ? ms : 1_000)
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
    const tooLarge = (why: string) => job.fail(new StreamSessionError('print-too-large', why), true)
    if (utf8Bytes(data) > TEXT_MESSAGE_BYTES) return hostile(`${String(m.type)} message over 48 KiB of UTF-8`)
    if (m.type === 'print-start') {
      if (job.scale) return hostile('a second print-start')
      const scale = printStart(m, job.n)
      if (!scale) return hostile('print-start does not match the job')
      // The whole job's decoded images, from the manifest sizes at the announced
      // scale, must fit PRINT_JOB_MAX_DECODED before any page is kept.
      if (printDecodedBound(this.manifest.units.list, job.from, job.from + job.n - 1, scale) > PRINT_JOB_MAX_DECODED) return tooLarge('the print job is over the decoded budget')
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
    let page: ReturnType<typeof printPage>
    try { page = printPage(payload, index, dims, job.scale) } catch (e) { return hostile((e as Error).message) }
    const { png } = page
    job.bytes += png.byteLength
    if (job.bytes > this.printMaxBlobBytes()) return job.fail(new StreamSessionError('print-too-large', 'the print job is over the blob ceiling'), true)
    // Defence in depth behind the print-start preflight: the checked raster sizes.
    job.decoded += page.w * page.h * 4
    if (job.decoded > PRINT_JOB_MAX_DECODED) return job.fail(new StreamSessionError('print-too-large', 'the print job is over the decoded budget'), true)
    job.pages.push({ unitIndex: index, widthPx: dims.w, heightPx: dims.h, scale: job.scale, png: new Blob([png as BlobPart], { type: 'image/png' }) })
    try { job.onProgress?.(job.pages.length, job.n) } catch (e) { job.fail(e, true) }
  }

  /** Sends the notes request under a fresh reqId (the retired one's partial
   * reassembly is dropped), unless `can.notes` went off meanwhile. A send that
   * throws fails it `policy-refused`. */
  private issueNotes(req: NotesReq) {
    if (req.reqId) { if (this.notesReqs.get(req.reqId) === req) this.notesReqs.delete(req.reqId); this.reasm.drop(req.reqId) }
    req.reqId = this.allocId(); req.sent = false; req.start = undefined; req.bodyCap = 0
    this.notesReqs.set(req.reqId, req)
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    if (!this.manifest.can.notes) return req.fail(new StreamSessionError('policy-refused', 'notes are off for this session'))
    try {
      this.send({ reqId: req.reqId, type: 'notes', token: this.token, unit: req.unit, scale: req.scale })
    } catch { return req.fail(new StreamSessionError('policy-refused', 'the notes request could not be sent')) }
    req.sent = true
    req.deadline = this.now() + REASSEMBLY_MS
    this.ensureSweeper()
  }
  /** The sent notes request `reqId` names on the current socket. */
  private liveNotes(reqId: unknown): NotesReq | undefined {
    const req = typeof reqId === 'number' ? this.notesReqs.get(reqId) : undefined
    return req?.sent ? req : undefined
  }
  /** The text messages of a notes request: a refusal (`error`, or the rate gate's `refused`), `notes-start`. */
  private onNotesText(req: NotesReq, m: Record<string, unknown>, data: string) {
    if (m.type === 'error' || m.type === 'refused') {
      // A RENEW race (stale jti, refused before anything was charged): re-issued once.
      if (m.code === 'session-rotated' && !req.start && !req.rotatedRetry) { req.rotatedRetry = true; return this.issueNotes(req) }
      const code = NOTES_REFUSALS.has(m.code as StreamErrorCode) ? m.code as StreamErrorCode : 'policy-refused'
      return req.fail(new StreamSessionError(code))
    }
    const hostile = (why: string) => req.fail(new StreamSessionError('policy-refused', why))
    if (utf8Bytes(data) > TEXT_MESSAGE_BYTES) return hostile(`${String(m.type)} message over 48 KiB of UTF-8`)
    if (m.type !== 'notes-start') return hostile(`unexpected ${String(m.type)} on a notes request`)
    if (req.start) return hostile('a second notes-start')
    const h = m.h
    if (m.scale !== req.scale || m.w !== NOTES_WIDTH_CSS * req.scale || !Number.isInteger(h) || (h as number) < 1 || (h as number) > NOTES_MAX_PX) return hostile('notes-start does not match the request')
    const w = m.w as number
    // §14/R10: the declared image's grant, now or never — its frame is already
    // on its way, and waiting in the serial inbox would block the messages whose leases free the room.
    const release = this.admission.tryAcquire(notesGrant(w, h as number))
    if (!release) return req.fail(new StreamSessionError('rate-limited', 'no admission room for the notes image'))
    req.hold(release)
    req.start = { w, h: h as number }
    req.bodyCap = notesBodyCap(w, h as number)
    // A declared image holds its grant: its frame must start arriving in time too.
    req.deadline = this.now() + REASSEMBLY_MS
  }
  /** The one NOTES_RASTER of a notes request: the unit asked, at the announced size. */
  private onNotesFrame(req: NotesReq, payload: Uint8Array) {
    const hostile = (why: string) => req.fail(new StreamSessionError('policy-refused', why))
    if (!req.start) return hostile('a notes image before notes-start')
    if (payload.length < 4) return hostile('truncated notes image')
    const unitIndex = new DataView(payload.buffer, payload.byteOffset, payload.byteLength).getUint32(0, true)
    if (unitIndex !== req.unit) return hostile(`notes image for unit ${unitIndex}, expected ${req.unit}`)
    let r: ReturnType<typeof decodeRasterBody>
    try { r = decodeRasterBody(payload.subarray(4), { maxRgbaBytes: UNIT_RASTER_DECODED_CAP }) } catch (e) { return hostile(`notes raster: ${(e as Error).message}`) }
    if (r.w !== req.start.w || r.h !== req.start.h) return hostile('notes image size is not the announced one')
    req.resolve(r)
  }

  /**
   * Sends the show request under a fresh reqId (the retired one's partial reassembly is dropped),
   * unless `can.present` went off meanwhile. `charged` false only for the stale-jti re-send, whose
   * log slot moves to the new reqId (R30). A charged send waits behind any queued one and for log room.
   */
  private issueShow(req: ShowReq, charged = true) {
    const prev = req.reqId
    if (prev) { if (this.shows.get(prev) === req) this.shows.delete(prev); this.reasm.drop(prev) }
    req.reqId = this.allocId(); req.sent = false; req.start = undefined; req.composited = undefined; req.received = []; req.bytes = 0; req.complete = false; req.admissionShort = false
    if (!charged) this.sendLog.transfer(prev, req.reqId)
    this.shows.set(req.reqId, req)
    if (this.state === 'resuming' || !this.sock || this.sock.readyState !== 1) return
    if (!this.manifest.can.present) return req.fail(new StreamSessionError('policy-refused', 'presenting is off for this session'))
    if (charged && (this.waiting.size || !this.hasRoom(req))) { this.waiting.add(req); return this.scheduleDrain() }
    this.transmitShow(req, charged)
  }
  private transmitShow(req: ShowReq, charged: boolean) {
    this.waiting.delete(req)
    try {
      this.send({ reqId: req.reqId, type: `show-${req.show}`, token: this.token, ...req.fields, scale: req.scale, motion: req.motion })
    } catch { return req.fail(new StreamSessionError('policy-refused', 'the show request could not be sent')) }
    req.sent = true
    req.deadline = this.now() + SHOW_REPLY_MS
    if (charged) {
      // R28/R29: one record per charged send; it holds its slot until the server's first reply.
      this.sendLog.record(req.reqId)
      try { req.onSend?.() } catch { /* a hook never breaks the send */ }
    }
    this.ensureSweeper()
  }
  /** The sent show request `reqId` names on the current socket. */
  private liveShow(reqId: unknown): ShowReq | undefined {
    const req = typeof reqId === 'number' ? this.shows.get(reqId) : undefined
    return req?.sent ? req : undefined
  }
  /**
   * Kind 6 is admitted only for a live show request with its start, short of its declared count, and
   * holding the grant of the message in flight: taken at a message's first fragment, now or never
   * (§14/R10: waiting in the serial inbox would block the messages whose completion frees the room).
   */
  private showAdmits(reqId: number): boolean {
    const req = this.liveShow(reqId)
    if (!req?.start || req.complete || req.received.length >= req.start.messages) return false
    if (req.grant) return true
    const grant = this.admission.tryAcquire(showGrant(req.start.decodedBytes))
    if (!grant) { req.admissionShort = true; return false }
    req.grant = grant
    return true
  }
  /** The text messages of a show request: a refusal, a `diag`, the `*-start`. */
  private onShowText(req: ShowReq, m: Record<string, unknown>, data: string) {
    if (m.type === 'error' || m.type === 'refused') {
      // A RENEW race (stale jti, refused before the rate gate): re-sent once, uncharged.
      if (m.code === 'session-rotated' && !req.start && !req.rotatedRetry) { req.rotatedRetry = true; return this.issueShow(req, false) }
      const code = SHOW_REFUSALS.has(m.code as StreamErrorCode) ? m.code as StreamErrorCode : 'policy-refused'
      return req.fail(new StreamSessionError(code))
    }
    const hostile = (why: string) => req.fail(new StreamSessionError('policy-refused', why))
    if (utf8Bytes(data) > SHOW_START_BYTES) return hostile(`${String(m.type)} message over 64 KiB of UTF-8`)
    // Task 7: diagnostics (forensic-unmarked, morph-composited, step-truncated) may precede the start.
    if (m.type === 'diag') return
    if (m.type !== `show-${req.show}-start`) return hostile(`unexpected ${String(m.type)} on a show-${req.show} request`)
    if (req.start) return hostile('a second start')
    if (m.unit !== req.unit || Object.entries(req.fields).some(([k, v]) => m[k] !== v)) return hostile(`show-${req.show}-start does not match the request`)
    // A closed key set per kind: a track or Morph start also declares its encoded bytes (D23), a plan start never.
    const windowed = req.show !== 'plan'
    const keys = new Set(['reqId', 'type', 'unit', 'messages', 'decodedBytes', ...Object.keys(req.fields), ...(windowed ? ['encodedBytes'] : [])])
    if (Object.keys(m).length !== keys.size || Object.keys(m).some((k) => !keys.has(k))) return hostile(`show-${req.show}-start has another key set`)
    const n = m.messages, decoded = m.decodedBytes, encoded = m.encodedBytes
    if (!Number.isInteger(n) || (n as number) < 1 || (n as number) > SHOW_MAX_MESSAGES[req.show]) return hostile('the declared message count is out of range')
    // §5 admission: the reply's decoded bytes held at once never exceed the show budget. A plan's are all its
    // assets together; a track's (and a Morph pair's) are its live D23 window, at most 8 decoded frames (plus the record).
    if (!Number.isInteger(decoded) || (decoded as number) < 1 || (decoded as number) > SHOW_PLAN_LIMITS.showDecoded) return hostile('the declared decoded bytes are out of range')
    // D23: a windowed reply's encoded PNGs stay retained (and charged) while it plays: within the show budget, and
    // within the kind-6 per-message encoded cap for every declared message.
    if (windowed && (!Number.isInteger(encoded) || (encoded as number) < 1 || (encoded as number) > SHOW_PLAN_LIMITS.showDecoded || (encoded as number) > (n as number) * CAPS[Kind.SHOW_PLAN]![0])) return hostile('the declared encoded bytes are out of range')
    // Nothing is reserved here: each message takes its grant at its first fragment (`showAdmits`).
    req.start = {
      messages: n as number, decodedBytes: decoded as number,
      encodedCap: showEncodedAllowance(decoded as number),
      bodyCap: Math.min(CAPS[Kind.SHOW_PLAN]![1], decoded as number),
      // A plan's retained bytes are bounded by its decoded total; a windowed reply's are exactly its declared encoded bytes.
      retainCap: windowed ? encoded as number : (decoded as number) + (n as number) * SHOW_MESSAGE_SLACK,
    }
    // A declared reply's messages must start arriving in time too.
    req.deadline = this.now() + REASSEMBLY_MS
  }
  /** One kind-6 message of a show request: in its kind's order, for the request's unit (a Morph record: the request's pair). */
  private onShowFrame(req: ShowReq, payload: Uint8Array) {
    const hostile = (why: string) => req.fail(new StreamSessionError('policy-refused', why))
    // Ruling 1: the message is reassembled, so its admission grant is done.
    if (req.grant) { const r = req.grant; req.grant = undefined; r() }
    const start = req.start
    if (!start || req.complete || req.received.length >= start.messages) return hostile('a show message past the declared reply')
    const i = req.received.length, tag = payload[0]
    const want = req.show === 'plan' ? (i === 0 ? 1 : i === 1 ? 3 : 2)
      : req.show === 'track' ? 4
      : i === 0 ? 5 : req.composited ? 4 : 2
    if (tag !== want) return hostile(`show message ${i} has tag ${String(tag)}, expected ${want}`)
    if (tag === 5) {
      let rec: ReturnType<typeof decodeShowMessage>
      try { rec = decodeShowMessage(payload, { unit: req.unit, maxDecoded: start.decodedBytes }) } catch (e) { return hostile(`Morph record: ${(e as Error).message}`) }
      if (rec.tag !== 5) return hostile('not a Morph record')
      const f = req.fields
      // Task 4 ruling: the record is the pair asked for, never another one.
      if (rec.morph.from !== f.from || rec.morph.to !== f.to || rec.morph.fromStep !== f.fromStep) return hostile('the Morph record is not the pair requested')
      if (rec.morph.composited && start.messages !== 1 + rec.morph.composited.count) return hostile('the composited Morph does not match the declared reply')
      req.composited = !!rec.morph.composited
    } else if (tag !== 1) {
      // An asset: `u32 LE unit` right after the tag must be the request's (the codec re-checks it all in decodeBundle).
      if (payload.length < 21) return hostile('a truncated show asset')
      const unit = new DataView(payload.buffer, payload.byteOffset, payload.byteLength).getUint32(1, true)
      if (unit !== req.unit) return hostile(`show asset for unit ${unit}, expected ${req.unit}`)
    }
    req.bytes += payload.length
    if (req.bytes > start.retainCap) return hostile('the show reply is over its declared bytes')
    // Ruling 5: a retained payload is charged to the show workspace at its encoded bytes.
    const lease = req.workspace.lease(payload.length, 'remote-show:message')
    if (!lease) return req.fail(new StreamSessionError('workspace-full', 'no workspace room for a show message'))
    req.leases.push(lease)
    req.received.push(payload)
    req.deadline = this.now() + REASSEMBLY_MS
    if (req.received.length < start.messages) return
    // Complete. It settles once the inbox has processed what was queued with it, so a surplus
    // message already received rejects the whole reply (D1 has no terminator: a later one is unsolicited).
    req.complete = true
    void this.inbox.then(() => { if (this.shows.get(req.reqId) === req) req.resolve(req.received) })
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

  private now() { return (this.o.now ?? (() => performance.now()))() }
  /** True once the session has ended: read synchronously at paint time (nothing paints after end). */
  isEnded(): boolean { return this.state === 'closed' }

  private allocId() { const id = this.nextReqId; this.nextReqId = id >= 0xfffffffe ? 1 : id + 1; return id }
  private send(m: object) { this.sock?.send(JSON.stringify(m)) }

  private connect(extra: object): Promise<Manifest> {
    const gen = ++this.gen
    this.reasm = new Reassembler({
      acceptedVersions: new Set([1]),
      now: () => this.now(),
      expects: (reqId, kind) => (kind === Kind.MANIFEST ? reqId === this.helloReqId && !!this.handshake
        // PRINT_RASTER only ever for the running, sent print job.
        : kind === Kind.PRINT_RASTER ? this.livePrint(reqId) !== undefined
        // NOTES_RASTER only ever for a live, sent notes request holding its declared grant.
        : kind === Kind.NOTES_RASTER ? this.liveNotes(reqId)?.start !== undefined
        // SHOW_PLAN only ever for a live, sent show request holding its declared grant, short of its declared count.
        : kind === Kind.SHOW_PLAN ? this.showAdmits(reqId)
        : this.units.has(reqId) && UNIT_KINDS.has(kind)),
      // A show message's declared total is held to its request's encoded allowance (the grant's encoded
      // part) as well as the kind cap, before any fragment of it is copied.
      encodedCap: (k, reqId) => Math.min(CAPS[k]?.[0] ?? 0, k === Kind.SHOW_PLAN ? this.liveShow(reqId)?.start?.encodedCap ?? 0 : Infinity),
      // A unit response may never decode past the bytes its request was admitted for.
      decodedCap: (k, reqId) => Math.min(CAPS[k]?.[1] ?? 0, k === Kind.MANIFEST ? Infinity
        : k === Kind.PRINT_RASTER ? this.livePrint(reqId)?.pageCap ?? 0
        : k === Kind.NOTES_RASTER ? this.liveNotes(reqId)?.bodyCap ?? 0
        : k === Kind.SHOW_PLAN ? this.liveShow(reqId)?.start?.bodyCap ?? 0
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
    // R29: any fragment for a unit send is evidence the server received it (also for an aborted request).
    if (data instanceof ArrayBuffer && data.byteLength >= HEADER_BYTES) this.unitAnswered(new DataView(data).getUint32(6, true))
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
    // R29: any reply starts its send's span — except a stale-token refusal about to be
    // re-sent (R30): the server never counted it, and its slot passes to the re-send.
    if (typeof m.reqId === 'number' && !this.rotatedResend(m)) this.unitAnswered(m.reqId)
    // R28: a show refusal (`error`) of a charged send blocks the log before the caller hears of it —
    // also when the request was aborted meanwhile (a notes refusal's reqId is never in the log).
    if (m.type === 'error' && m.code === 'rate-limited' && typeof m.reqId === 'number' && this.sendLog.has(m.reqId)) this.sendLog.refused()
    if (m.type === 'renew' && typeof m.token === 'string' && typeof m.jti === 'string' && jtiOf(m.token) === m.jti) {
      this.send({ reqId: this.allocId(), type: 'renew-ack', jti: m.jti })
      this.token = m.token
      return
    }
    const job = this.livePrint(m.reqId)
    if (job) return this.onPrintText(job, m, data)
    const notes = this.liveNotes(m.reqId)
    if (notes) return this.onNotesText(notes, m, data)
    const show = this.liveShow(m.reqId)
    if (show) return this.onShowText(show, m, data)
    const text = typeof m.reqId === 'number' ? this.texts.get(m.reqId) : undefined
    // Only text-service answers are held to 48 KiB (refused/diag are control messages).
    if (text) return this.onTextAnswer(text, m, TEXT_ANSWER_TYPES.has(m.type as string) ? utf8Bytes(data) : 0)
    if (m.type === 'refused' && typeof m.reqId === 'number') {
      const req = this.units.get(m.reqId)
      // R28: the log blocks every charged unit send (any caller) before the caller
      // hears of it — also when the refused request was aborted meanwhile.
      if (m.code === 'rate-limited' && (this.sendLog.has(m.reqId) || req)) this.sendLog.refused()
      if (!req) return
      if (m.code === 'session-rotated' && !req.rotatedRetry) { req.rotatedRetry = true; return this.issue(req, false) }
      const code = m.code === 'rate-limited' || m.code === 'session-rotated' ? m.code : 'policy-refused'
      req.reject(new StreamSessionError(code))
    }
  }

  private async onFrame(res: FrameResult) {
    if (!res.ok) {
      // Fail closed: any rejected frame on the print job fails the job.
      const job = this.livePrint(res.reqId)
      if (job) return job.fail(new StreamSessionError('policy-refused', `print frame rejected: ${res.reason}`), true)
      const notes = this.liveNotes(res.reqId)
      if (notes) return notes.fail(new StreamSessionError('policy-refused', `notes frame rejected: ${res.reason}`))
      // Fail closed: a kind-6 message before the start or past the declared count, over a cap, or malformed.
      const show = this.liveShow(res.reqId)
      if (show) return show.fail(show.admissionShort ? new StreamSessionError('admission', 'no admission room for a show message') : new StreamSessionError('policy-refused', `show frame rejected: ${res.reason}`))
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
    if (res.kind === Kind.NOTES_RASTER) {
      // A message that completed after its request settled is dropped.
      const notes = this.liveNotes(res.reqId)
      return notes && this.onNotesFrame(notes, res.payload)
    }
    if (res.kind === Kind.SHOW_PLAN) {
      // A message that completed after its request settled is dropped.
      const show = this.liveShow(res.reqId)
      return show && this.onShowFrame(show, res.payload)
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
    // Likewise a sent notes request (the server audits each one): never re-issued.
    for (const req of [...this.notesReqs.values()]) if (req.sent) req.fail(new StreamSessionError('session-rotated', 'the session resumed with a notes request in flight'))
    // Likewise a sent show request (the server charged and audited it; a re-issue is a second charge).
    for (const req of [...this.shows.values()]) if (req.sent) req.fail(new StreamSessionError('session-rotated', 'the session resumed with a show request in flight'))
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
    // R29: a send still unanswered from the old socket is taken as received now (the resume reply).
    this.sendLog.settleAll()
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
    // Notes requested while resuming go out now, if the new manifest still offers notes.
    for (const req of [...this.notesReqs.values()]) if (!req.sent) this.issueNotes(req)
    // Show requests asked for while resuming (or waiting for log room) go out now, through the log.
    for (const req of [...this.shows.values()]) if (!req.sent) this.issueShow(req)
  }

  private ensureSweeper() {
    if (this.sweeper) return
    this.sweeper = setInterval(() => {
      for (const r of this.reasm.sweep()) {
        const job = this.livePrint(r.reqId)
        if (job) { job.fail(new StreamSessionError('policy-refused', 'a print page timed out in reassembly'), true); continue }
        const notes = this.liveNotes(r.reqId)
        if (notes) { notes.fail(new StreamSessionError('policy-refused', 'a notes image timed out in reassembly')); continue }
        const show = this.liveShow(r.reqId)
        if (show) { show.fail(new StreamSessionError('policy-refused', 'a show message timed out in reassembly')); continue }
        const req = r.reqId !== undefined ? this.units.get(r.reqId) : undefined
        if (req) this.retryOrFail(req)
      }
      // A notes request with nothing in reassembly (no declaration, or a declaration and silence) expires on its own deadline.
      const now = this.now()
      for (const req of [...this.notesReqs.values()]) {
        if (req.sent && now >= req.deadline && !this.reasm.has(req.reqId)) req.fail(new StreamSessionError('policy-refused', 'the notes request timed out'))
      }
      for (const req of [...this.shows.values()]) {
        if (req.sent && !req.complete && now >= req.deadline && !this.reasm.has(req.reqId)) req.fail(new StreamSessionError('policy-refused', 'the show request timed out'))
      }
      if (!this.units.size && !this.printJob && !this.notesReqs.size && !this.shows.size && this.sweeper) { clearInterval(this.sweeper); this.sweeper = undefined }
    }, 1000)
  }

  private end(error: StreamSessionError | undefined, closeCode?: number, reason?: string) {
    if (this.state === 'closed') return
    this.gen++
    this.failure = error
    if (this.sweeper) { clearInterval(this.sweeper); this.sweeper = undefined }
    if (this.drainTimer) { clearTimeout(this.drainTimer); this.drainTimer = undefined }
    const sock = this.sock
    this.sock = undefined
    if (sock && closeCode !== undefined && sock.readyState < 2) sock.close(closeCode, reason)
    const err = error ?? abortError()
    if (this.handshake) { this.handshake.reject(err); this.handshake = undefined }
    for (const req of [...this.units.values()]) req.reject(err)
    for (const req of [...this.texts.values()]) req.reject(err)
    this.printJob?.fail(err, false)
    for (const req of [...this.notesReqs.values()]) req.fail(err)
    for (const req of [...this.shows.values()]) req.fail(err)
    this.admission.drain(err)
    this.setState('closed', error)
    this.settleEnded(error)
  }
}
