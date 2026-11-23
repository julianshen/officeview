/**
 * The only endpoint (§9.1): WS upgrade with an origin allowlist, the §6.2
 * state machine (HELLO → MANIFEST → READY → STREAM, RENEW, RESUME), per-op
 * token + capability checks, §14 rate gates, and the §6.2 rule-8 close-code
 * registry. No HTTP document route exists: every non-upgrade request is 404.
 */
import { createHash, randomBytes } from 'node:crypto'
import { open as openFile } from 'node:fs/promises'
import { gzipSync } from 'node:zlib'
import type { Server, ServerWebSocket } from 'bun'
import { encodeMessage } from '../../src/stream/frame'
import { encodeUnit } from '../../src/stream/unit'
import { encodeRasterBody } from '../../src/stream/raster'
import { AssetTooLarge, encodeAssets } from '../../src/stream/asset-wire'
import { ASSET_ALLOWANCE } from '../../src/stream/remote/reservation'
import { encodeManifest, MEDIA_AUDIO_CAPS, type Manifest } from '../../src/stream/manifest'
import { bindMedia, mediaIdentity, sessionMediaId } from './media/bind'
import type { FfmpegInfo } from './media/ffmpeg'
import { mediaJobsFor, type MediaItem } from './media/jobs'
import type { SessionMediaQueue } from './media/queue'
import type { MediaServices } from './media/services'
import { MEDIA_WIRE_LIMITS, parseMediaCursor, parseMediaEnd, parseMediaRequest, parseMediaShow, type MediaRefused } from '../../src/stream/media-wire'
import { MEDIA_ID } from '../../src/stream/show-plan'
import type { MediaRefusal } from '../../src/stream/media-types'
import { AuditTable, ForensicExhausted, msgHex, type Issued } from './audit'
import { CODEC_ID } from './mark/params'
import { encodeEstimate } from './mark/mark-raster'
import { deriveBasisKey } from './mark/basis'
import { LeaseFull, ScratchLease } from './scratch'
import { createRenderCache, DocumentLedger, renderKey, viewStateHash, type RenderKeyFields } from './cache'
import { loadPack, requiredFacesOf } from './fonts'
import { canNotes, effectivePolicy, forensicNotices, resumeInvariant, type AuthDecision, type AuthRequest, type Authorize, type EffectivePolicy, type ResolveDocument } from './hooks'
import { docRefOf, type Keys } from './keys'
import { admit } from './policy'
import { AdmissionRefused, RenderFailed, RenderTimeout, type Pool } from './pool'
import type { SheetTextReply } from './sheet/text-ops'
import { showBuildGrant, showMorphGrant } from './show/grant'
import { encodeShowMessage, PlanRejected, SHOW_TRACK_WINDOW, type PlanMediaRecord, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import type { TestOnlyOptions } from './config'
import type { ShowDeliveryRecord } from './audit-store'
import { encodeNotesBody, NOTES_ENCODED_CAP, notesBasisLabel } from './notes'
import { encodePrintBody, planPrint, PRINT_ENCODED_CAP, PRINT_SEND_HIGH_WATER, printArgs, printSizeOk, runPrint, type PrintJob, type PrintSlot } from './print'
import { IpRates, SessionRates, type RateOverrides } from './rate'
import { layoutFingerprint, SCALE_BUCKETS, type MarkSpec, type MediaUnitRefs, type NotesPlan, type NotesRaster, type RenderedUnit, type UnitSize } from './render'
import { SessionTable, type Session } from './sessionTable'
import { a11yArgs, rangeArg, runA11y, runCopy, runSearch, runSelect, searchArgs, type SearchRun, type TextSlot } from './text/ops'
import { issueToken, newSid, verifyToken } from './tokens'
import { parseFeatures, sheetModeOf, type SheetMode } from './sheet/mode'
import { sheetClockAt, type SheetClock } from './sheet/clock'
import { encodeSheetLayout, encodeTileBody, fitArgs, LayoutTooLarge, MAX_SHEETS, resizeArgs, SHEET_NAME_CHARS, SHEET_VIEW_FEATURE, TILE_BATCH_MAX, TILE_DPRS, TILE_PALETTES, TILE_PX, TILE_ZOOM_STEPS, type StreamSheetLayout } from '../../src/stream/sheet-wire'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { TILE_BYTES_ESTIMATE, tileBasisLabel, type RenderedTile } from './sheet/tile'
import { runTiles, TileSlots, tilesArgs, tilesInGrid, type SheetExtent, type TileBatch, type TileCoord, type TileRefusalCode, type TilesArgs } from './sheet/tiles-op'
import { sessionLayout, type ServerSheetLayout } from './sheet/layout'
import { SessionViews, ViewStateLedger, type WireSizes } from './sheet/view-state'

/** §6.2 rule 8 — the complete registry. 4007 is client-only and never sent here. */
export const CLOSE = {
  'unsupported-version': 4000, 'unsupported-format': 4001, 'ready-timeout': 4002, 'auth-failed': 4003,
  'session-expired': 4004, 'revoked': 4005, 'layout-changed': 4006, 'rate-limited': 4008, 'policy-refused': 4009,
} as const
export type CloseName = keyof typeof CLOSE
export type RefusalCode = 'rate-limited' | 'policy-refused' | 'session-rotated' | 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted' | 'print-too-large' | 'print-quota-exhausted' | 'unsupported-op' | 'stale-view' | 'layout-too-large' | 'hidden-target' | 'coverage-unknown' | 'view-state-limit'
/** A notes refusal (Stage 3c), sent as `{reqId, type: 'error', code}`; `session-rotated` is the token preamble's retry signal. */
export type NotesErrorCode = 'policy-refused' | 'invalid-unit' | 'rate-limited' | 'session-rotated'
/**
 * A show op refusal (piece 4 Task 7, D1), sent as `{reqId, type: 'error', code}`. Every one plays the slide's
 * piece-3 still on the client (a Morph: the fade look-alike), never a plan with content withheld:
 * - `invalid-unit`: hidden, out of range, not animated (D5), a step or `fromStep` the slide does not have, a Morph
 *   pair whose `to` is not entered by Morph;
 * - `show-budget`, `show-limits`, `show-plan-timeout`: the build's own refusals (D3, D10, §5 admission);
 * - `still` (controller ruling): the classifier cannot prove the slide (`Classification.fallback`, a trigger that
 *   can move a layer, a withdrawn interactive slide): the client plays the piece-3 still;
 * - `morph-fallback` (`show-morph` only): `planMorph` refused (past 256 objects, its layers do not fit);
 * - `rate-limited`, `policy-refused`, `session-rotated`: as for notes.
 */
export type ShowErrorCode = NotesErrorCode | 'show-budget' | 'show-limits' | 'show-plan-timeout' | 'still' | 'morph-fallback'
/**
 * D23 (final review finding 1): the decoded bytes a keyframe track holds at once — at most `SHOW_TRACK_WINDOW`
 * frames of the largest frame's `w·h·4` — which its start declares, not Σ its frames (never decoded together).
 */
export const trackWindowBytes = (frames: ReadonlyArray<{ w: number; h: number }>) =>
  Math.min(frames.length, SHOW_TRACK_WINDOW) * frames.reduce((m, f) => Math.max(m, f.w * f.h * 4), 0)

/** The most click steps a plan carries (D3): a track's step is below it, a Morph's `fromStep` at most it. */
const SHOW_PLAN_LIMITS_STEPS = 64
const KIND = { MANIFEST: 1, ASSET_DEFINE: 2, UNIT_RASTER: 3, UNIT_COMMANDS: 4, XLSX_TILE: 5, SHOW_PLAN: 6, NOTES_RASTER: 7, PRINT_RASTER: 8, MEDIA_CHUNK: 9 } as const
/** Admission charge of a notes plan (a measuring context and the notes part's parse). */
const NOTES_PLAN_BYTES = 4 * 1024 * 1024
/** R6: the ops an XLSX tiles session refuses `unsupported-op`. */
const TILES_REFUSED: ReadonlySet<string> = new Set(['unit', 'search', 'select', 'copy', 'a11y'])
/** Stage 4b (P1): the ops served only to a `sheet-view@1` session; every other session refuses them `unsupported-op`. */
const SHEET_VIEW_OPS: ReadonlySet<string> = new Set(['sheet-find', 'sheet-copy', 'sheet-a11y', 'sheet-fit', 'sheet-resize'])
/** Stage 4b: the SHEET_VIEW_OPS that ride the `'text'` rate class (the rest ride `'sheet'`). */
const SHEET_TEXT_OPS: ReadonlySet<string> = new Set(['sheet-find', 'sheet-copy', 'sheet-a11y'])
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
  /** Piece 4 (spec §7): `strictMarks` plays every moving step `composited` (`config.animation`). */
  animation?: { strictMarks?: boolean }
  /** TEST-ONLY (`config.ts` `TestOnlyOptions`): never read from an `AuthDecision`. */
  testOnly?: TestOnlyOptions
  /**
   * Piece 5 (D18): audio media in shows. `enabled` defaults to `true` (Task 11: the client plays it); `can.media` is
   * `enabled && info !== undefined` (ffmpeg probed: no `info` is media unavailable, fail closed). Off
   * (`enabled: false`), shows behave exactly as in piece 4.
   */
  /** TEST-ONLY: awaited inside a media delivery, after the file is opened and checked, and after each chunk read. */
  mediaTestHooks?: { afterOpen?(): void | Promise<void>; afterRead?(chunk: number): void | Promise<void> }
  /** Stage 4a (R9): the session clock's zone when the decision's `policy.timeZone` is absent. Default `'UTC'`; a zone
   * `Intl` rejects throws at construction. */
  sheetTimeZone?: string
  /** TEST and measure only (R3): a `'message'` or `'tile'` override replaces that class's windows, its burst included. */
  rates?: RateOverrides
  /** TEST and measure only (R3, T8): per delivered tile, `t0` is the instant its batch passed validation, the token
   * and the `'tile'` rate check, and `t1` the instant its last fragment was passed to `sendBinary` (both
   * `performance.now()`). Unset, no clock is read. */
  onTileTiming?(t: { t0: number; t1: number; zoomStep: number; dpr: number }): void
  media?: { enabled?: boolean; info?: FfmpegInfo; /** Task 9: the process's media services (queues, the shared cache); without them `media-show` and `media-request` answer `media-unavailable`. */ services?: MediaServices }
  /** Stage 4b (P12): the process's view-state ledger (default 16 MiB; tests inject a smaller one). */
  viewStateLedger?: ViewStateLedger
}

/** Piece 5 (Task 9): one audio use of the session on one slide: its encode item (`mediaJobsFor`), or the refusal the
 * slide's plan record ships (probe, link, or a chain on a video end: `slideAudioUses`). Server-only. */
interface MediaUse { unit: number; refused?: MediaRefusal; item?: MediaItem }
/**
 * Piece 5 (Task 9): the session's media state on its `sessionDoc` entry.
 * - `queue`: created at `media-show`, ended at `media-end`, on every socket close, expiry and revocation.
 * - `uses`: mediaId → slide → use (fix round 1: a spid-less `audio|k` is one mediaId on every slide that uses its part,
 *   and may be refused on one and play on another). Filled as the queue loads the slides' items (`started` settles
 *   once it has); every id the session's plans can carry is here, and nothing else.
 * - `refs`: each slide's `pool.mediaRefs`, read once per session (a re-show after `media-end` or a resume reuses it).
 * - `cursor`: the slide of the last `media-show` or `media-cursor` (a `media-request` without `unit` is for it).
 * - `wakes`: the high-water wait of every delivery in progress, woken by `endMedia` and the document's revocation.
 */
interface MediaSlot {
  queue?: SessionMediaQueue; started?: Promise<void>; uses: Map<string, Map<number, MediaUse>>
  refs: Map<number, MediaUnitRefs>; cursor?: number; wakes: Set<() => void>; revocationHooked?: boolean
  /**
   * Task 11b fix round (controller ruling): per unit, the media answers of the last plan bound for this session
   * (mediaId → its plan refusal, `undefined` when it plays). A use is delivered only when a plan sent for its unit
   * ships it playable; a unit no plan was sent for delivers nothing (`media-unavailable`).
   */
  plans: Map<number, Map<string, PlanMediaRecord['refused']>>
}

/** Stage 4a: a sid's tile state (`sessionDoc`). Resume and teardown stop every batch. */
interface TileState { batches: Map<number, TileBatch>; slots: TileSlots; views: Set<string> }

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
  /** Aborted when the socket closes: its queued show builds and Layer-1 encodes never start (PR #78 round 4). */
  gone: AbortController
}
type WS = ServerWebSocket<Conn>

/** `hidden`: unit indices of hidden slides (pptx under 'full'); the show skips them and notes refuse them. */
/** `anim` (D5): shown units with timing, a transition or both; `morph` (Task 7): those entered by Morph. */
/** `media` (piece 5, D19): shown units whose plan ships an audio record or a transition sound. */
interface DocState { docRef: string; name: string; bytes: Uint8Array; units: UnitSize[]; hidden: ReadonlySet<number>; anim: ReadonlySet<number>; morph: ReadonlySet<number>; media: ReadonlySet<number>; fingerprint: string; contentPolicyId: string; format: 'docx' | 'xlsx' | 'pptx'; refs: number; /** XLSX: the projected (visible) sheet names in tab order (Stage 4a). */ sheets?: readonly string[] }

const isU32 = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0 && (n as number) <= 0xffffffff

export function createStreamServer(o: StreamServerOptions) {
  const now = o.now ?? Date.now
  const sheetTimeZone = o.sheetTimeZone ?? 'UTC'
  sheetClockAt(0, sheetTimeZone)
  /** R9: the zone a decision resolves to (validated by `sheetClockAt`). */
  const zoneOf = (d: AuthDecision): unknown => d.policy?.timeZone ?? sheetTimeZone
  const nowSec = () => Math.floor(now() / 1000)
  const ttl = o.tokenTtlSec ?? 300, lead = o.renewLeadSec ?? 5
  const readyMs = o.readyMs ?? 2000, helloMs = o.helloMs ?? 5000, graceMs = o.detachGraceMs ?? 30_000
  /** sid → reaper of a detached session. */
  const reapers = new Map<string, ReturnType<typeof setTimeout>>()
  const fingerprintOf = o.fingerprint ?? layoutFingerprint
  /** Piece 5 (D18): `can.media`, for every session of this server. */
  const canMedia = (o.media?.enabled ?? true) && o.media?.info !== undefined
  /** Media deliveries in progress (ok answers, until their pin is released): `stats().media`. */
  let mediaDeliveries = 0
  /** Task 9: `media-show` / `media-request` are served only with `can.media` and the media services. */
  const services = canMedia ? o.media?.services : undefined
  const table = new SessionTable()
  const audit = o.audit ?? new AuditTable()
  const ipRates = new IpRates()
  const docs = new Map<string, DocState>()          // docRef → state
  /** Per-sid session entry. Its text state (in-flight search/select, the copy
   * queue and the quota counter `copiedParagraphs`) survives resume and is
   * gone on teardown. So does its print state (`printedPages`, the running job). */
  const sessionDoc = new Map<string, TextSlot & PrintSlot & { doc: DocState; policy: EffectivePolicy; decision: AuthDecision; tokenExp: number; /** Layer-1 issuance, fixed for the sid's life (a resume keeps it). */ forensic?: Issued; /** Piece 5 (D2): `sessionMediaKey`, 32 random bytes fixed for the sid's life (a resume keeps it). */ mediaKey: Uint8Array; /** Piece 5 (Task 9). */ media: MediaSlot; /** The sid's in-flight search (R6: one per sid); survives resume. */ search?: SearchRun; /** Stage 4a: the HELLO's validated `features`, fixed for the sid (a resume must present the same list). */ features: readonly string[]; /** Stage 4a: fixed at open. */ sheetMode: SheetMode; /** Stage 4b (P1): true iff the document is in tiles mode and `features` hold `sheet-view@1`; fixed for the sid (a resume presents the same features). */ sheetView: boolean; /** Stage 4a (R9): fixed at open; survives resume. */ clock: SheetClock; /** Stage 4a (R4): the sid's `tiles` batches in progress (by reqId), its tile slots, and the `(sheetId, viewGen)` views already recorded in `sheet_views`. */ tiles: TileState; /** Stage 4b (S4 §6.3): the sid's per-sheet sizes and `viewGen`; survives resume, disposed at teardown. */ views: SessionViews }>()
  const sockets = new Set<WS>()
  /** Per-sid rate state survives resume (it belongs to the session, not the socket). */
  const sessionRates = new Map<string, SessionRates>()
  /** Units and (Stage 4a) canonical sheet tiles, under one 512 MiB ceiling; a tile key never equals a unit key (its `tile` field is hashed). */
  const renderCache = createRenderCache<RenderedUnit | RenderedTile>()
  /** Stage 4a (F8): sheet extents by `docRef|contentPolicyId|fingerprint|sheetIndex|day|timeZone`, from the sheet's
   * layout (no cell data). Bounded; the oldest entry goes first. */
  const extents = new Map<string, Promise<SheetExtent>>()
  const EXTENTS_MAX = 1024
  /** Stage 4b: identity layouts (with the server-only `coverageUnknown`) by the extent's base key, for the extents of
   * sized views and `sheet-resize`'s checks. Bounded; the oldest entry goes first. */
  const baseLayouts = new Map<string, Promise<ServerSheetLayout>>()
  const BASE_LAYOUTS_MAX = 16
  /** Stage 4b (P12): the process cap on view-state bytes. */
  const viewLedger = o.viewStateLedger ?? new ViewStateLedger()
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
    // D1 (fix round 1): a server-initiated close ends the session's media now, not at the peer's close handshake.
    const sid = ws.data.sid
    if (sid && table.get(sid)?.activeSocket === (ws as unknown)) endMedia(sessionDoc.get(sid))
    ws.data.phase = 'closed'
    ws.data.gone.abort()
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
  const ENCODED: Record<number, number> = { [KIND.MANIFEST]: 256 << 10, [KIND.ASSET_DEFINE]: 1 << 20, [KIND.UNIT_RASTER]: 8 << 20, [KIND.UNIT_COMMANDS]: 8 << 20, [KIND.XLSX_TILE]: 8 << 20, [KIND.SHOW_PLAN]: 8 << 20, [KIND.NOTES_RASTER]: NOTES_ENCODED_CAP, [KIND.PRINT_RASTER]: PRINT_ENCODED_CAP, [KIND.MEDIA_CHUNK]: 4 << 20 }
  const frame = (kind: number, reqId: number, payload: Uint8Array): Uint8Array[] | undefined => {
    const z = new Uint8Array(gzipSync(payload, { level: 6 }))
    if (z.length + 4 > ENCODED[kind]!) return undefined
    return encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: () => z, fragmentBytes: FRAGMENT })
  }
  /** D13, D14: one MEDIA_CHUNK, framed stored (already-compressed audio), one fragment. */
  const frameChunk = (reqId: number, payload: Uint8Array): Uint8Array[] | undefined => {
    if (payload.length + 4 > ENCODED[KIND.MEDIA_CHUNK]!) return undefined
    return encodeMessage({ ver: 1, kind: KIND.MEDIA_CHUNK as never, reqId, payload, compress: () => { throw new Error('a stored message is never compressed') }, stored: true, fragmentBytes: FRAGMENT })
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
    if (d) { d.refs++; return }
    docs.set(docRef, { docRef, refs: 1 } as DocState)
    // A document opened again after its close (the same docRef): its media is served afresh (Task 9).
    o.media?.services?.cache.reopen(docRef)
  }
  /** The last release closes the document: the pool forgets it, and its media entries are deleted at once (D12, D23). */
  function releaseDoc(docRef: string): void {
    const d = docs.get(docRef)
    if (d && --d.refs <= 0) { docs.delete(docRef); o.pool.forget(docRef); void o.media?.services?.cache.dropDocument(docRef).catch(() => {}) }
  }
  /** D1: `media-end`, a socket close, expiry and revocation all end the session's queue (`cache.dropSession`). */
  function endMedia(sd: { media: MediaSlot } | undefined): void {
    if (!sd) return
    // Deliveries parked in the high-water wait exit at once (fix round 1); their alive check fails.
    const q = sd.media.queue
    sd.media.queue = undefined; sd.media.started = undefined
    for (const w of [...sd.media.wakes]) w()
    if (q) void q.end().catch(() => {})
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
    // validHello has checked it.
    const features = parseFeatures(hello.features) as readonly string[]
    if (resume !== undefined) return resumeSession(ws, req, resume, reqId, features)
    const decision = await o.authorize(req)
    if (gone(ws)) return
    if (!decision) return close(ws, 'auth-failed')
    const resolved = await o.resolveDocument(req, decision)
    if (gone(ws)) return
    if (!resolved) return close(ws, 'auth-failed')
    const admitted = admit(resolved.bytes)
    if (!admitted.ok) return close(ws, 'unsupported-format')
    const policy = effectivePolicy(decision)
    // R9: the session clock, before anything is charged or issued; an invalid zone refuses the open.
    let clock: SheetClock
    try { clock = sheetClockAt(now(), zoneOf(decision) as string) } catch { return close(ws, 'policy-refused') }
    const requestedMode = sheetModeOf(admitted.format, features)
    const docRef = await docRefOf(o.keys.kDoc, resolved.bytes)
    if (gone(ws)) return
    // Charge the document BEFORE parsing; one that cannot fit is refused at admission.
    if (!ledger.touch(docRef, policy.contentPolicy, resolved.bytes.length)) return close(ws, 'unsupported-format')
    acquireDoc(docRef)
    let committed = false
    try {
      let opened
      try { opened = admitted.format === 'xlsx' ? await o.pool.open(docRef, resolved.bytes, policy.contentPolicy, { sheetMode: requestedMode }) : await o.pool.open(docRef, resolved.bytes, policy.contentPolicy) } catch (e) {
        return close(ws, (e as RenderFailed).refusal === 'unsupported-format' ? 'unsupported-format' : 'auth-failed')
      }
      if (gone(ws)) return
      // Fix round 1: a workbook a tiles manifest cannot list (over 256 sheets, or a name over 31 characters) is served
      // in units mode; its paintables build on first units use (R5).
      const sheets = opened.sheets ?? []
      const sheetMode: SheetMode = requestedMode === 'tiles' && (sheets.length > MAX_SHEETS || sheets.some((n) => n.length > SHEET_NAME_CHARS)) ? 'units' : requestedMode
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
        audit.record(sid, { ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), ip: ws.data.ip, startedAt, keyGeneration: o.keys.keyGeneration, policyId: opened.contentPolicyId, markSettings: { visible: policy.visibleMark, forensicRequested: policy.forensicRequested }, relaxed: policy.relaxed, services: { copy: policy.copy, textA11y: policy.textA11y, print: policy.print, notes: canNotes(policy), media: canMedia }, forensic: forensic ? { codec: CODEC_ID, token24: forensic.token24, salt8: forensic.salt8, msg40: msgHex(forensic.msg40) } : null, docRef, contentPolicy: policy.contentPolicy, visibleMark: markOf(policy, decision), sheetMode })
      } catch { return close(ws, 'policy-refused') }
      table.create({ sid, docRef, ...(decision.viewerId ? { viewerId: decision.viewerId } : {}), role: policy.role, effectivePolicy: policy, activeSocket: ws as never, jti: token.jti, startedAt })
      sessionDoc.set(sid, { doc: { docRef, name: resolved.name ?? 'document', bytes: resolved.bytes, units: opened.units, hidden: new Set(opened.hidden), anim: new Set(opened.anim.filter((i) => !opened.hidden.includes(i))), morph: new Set(opened.morph.filter((i) => !opened.hidden.includes(i))), media: new Set((opened.media ?? []).filter((i) => !opened.hidden.includes(i))), fingerprint: opened.layoutFingerprint, contentPolicyId: opened.contentPolicyId, format: opened.format, refs: 0, ...(opened.sheets ? { sheets: opened.sheets } : {}) }, policy, decision, tokenExp: token.exp, features, sheetMode, sheetView: sheetMode === 'tiles' && features.includes(SHEET_VIEW_FEATURE), clock, tiles: { batches: new Map(), slots: new TileSlots(o.pool.admission.caps.perSessionOutstanding), views: new Set() }, views: new SessionViews(viewLedger), copiedParagraphs: 0, printedPages: 0, mediaKey: new Uint8Array(randomBytes(32)), media: { uses: new Map(), refs: new Map(), wakes: new Set(), plans: new Map() }, ...(forensic ? { forensic } : {}) })
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
  async function resumeSession(ws: WS, req: AuthRequest, resume: { sid?: unknown; resumeProof?: unknown }, reqId: number, features: readonly string[]) {
    if (typeof resume.sid !== 'string' || typeof resume.resumeProof !== 'string') return close(ws, 'auth-failed')
    const s = table.get(resume.sid), sd = sessionDoc.get(resume.sid)
    if (!s || !sd) return close(ws, 'auth-failed')
    const generation = s.jti
    const v = await verifyToken(o.keys.kToken, resume.resumeProof, { now: nowSec(), docRef: s.docRef, sid: s.sid })
    if (gone(ws)) return
    const tokenProof = v.ok && ['current', 'rotating'].includes(table.acceptJti(s.sid, v.claims.jti))
    const d = await o.authorize({ ...req, resume: { sid: s.sid }, ...(tokenProof ? {} : { token: resume.resumeProof }) })
    if (gone(ws)) return
    // R3: copy, textA11y, the print caps (Stage 3b P4) and notes (Stage 3c) may change on resume; every other field, role included, is exact-match.
    const policy = d ? effectivePolicy(d) : undefined
    const ok = !!d && !!policy && (tokenProof || d.resumeGranted === true) && d.viewerId === s.viewerId && JSON.stringify(resumeInvariant(policy)) === JSON.stringify(resumeInvariant(s.effectivePolicy))
      // Stage 4a: the same features (so the same sheet mode) and the same clock zone (R9); the day is kept.
      && JSON.stringify(features) === JSON.stringify(sd.features) && zoneOf(d) === sd.clock.timeZone
    if (!ok || table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    const token = await issueToken(o.keys.kToken, { docRef: s.docRef, sid: s.sid, ttlSec: ttl, now: nowSec() })
    if (gone(ws)) return
    if (table.get(s.sid) !== s || s.jti !== generation) return close(ws, 'auth-failed')
    // A change of the mutable text services is durable before it takes effect
    // (the open row stays write-once); a failed write refuses the resume.
    if (JSON.stringify([policy!.copy, policy!.textA11y, policy!.print, canNotes(policy!)]) !== JSON.stringify([sd.policy.copy, sd.policy.textA11y, sd.policy.print, canNotes(sd.policy)])) {
      try { audit.recordServices(s.sid, { copy: policy!.copy, textA11y: policy!.textA11y, print: policy!.print, notes: canNotes(policy!), media: canMedia }) } catch { return close(ws, 'policy-refused') }
    }
    const old = s.activeSocket as unknown as WS | undefined
    if (old && old !== ws && old.data) { old.data.phase = 'closed'; for (const t of old.data.timers) clearTimeout(t); wakeDrain(old) }
    // D1: the superseded socket's close is an implicit media-end (its deliveries stop on their own alive check).
    if (old !== ws) endMedia(sd)
    // F21: a `'sheet'` op of the old socket is over; the re-issued op sees a clean flag.
    const rates = sessionRates.get(s.sid); if (rates) rates.sheetBusy = false
    // A job on the old socket cannot continue on the new one.
    if (sd.print) { sd.print.stopped = true; sd.print.wake?.() }
    stopTiles(sd)
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

  /** Stage 4a: no tile of the sid's batches starts after this (resume, teardown); tiles in flight are sent only while their socket owns the sid. */
  function stopTiles(sd: { tiles: TileState }) { for (const b of sd.tiles.batches.values()) b.stopped = true }

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
    if (sd) stopTiles(sd)
    sd?.views.dispose()
    endMedia(sd)
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
      // D19: `media` only with `can.media` (D18: off, the manifest is piece 4's).
      units: { count: sd.doc.units.length, list: sd.doc.units.map((u, i) => sd.doc.hidden.has(i) ? { ...u, hidden: true as const } : { ...u, ...(sd.doc.anim.has(i) ? { anim: true as const } : {}), ...(canMedia && sd.doc.media.has(i) ? { media: true as const } : {}) }) },
      layoutFingerprint: sd.doc.fingerprint,
      session: { sid, token, expiresSec: ttl, resume: 'session' },
      policy: { contentPolicy: p.contentPolicy, protection: p.projection, semantic: 'visible-only', contentPolicyId: sd.doc.contentPolicyId, viewerMarkId: p.visibleMark },
      // Links land later in Stage 3.
      can: { search: true, copy: p.copy !== false, print: p.print !== false, textA11y: p.textA11y !== 'off', present: sd.doc.format === 'pptx', notes: canNotes(p), links: false, media: canMedia },
      caps: { copy: p.copy === false ? null : { ...p.copy }, a11y: p.textA11y, print: p.print === false ? null : { ...p.print }, ...(canMedia ? { media: { audio: { ...MEDIA_AUDIO_CAPS } } } : {}) },
      // D18: off, `caps` stays piece 4's (absent `caps.media` validates as null, media off).
      recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [...SCALE_BUCKETS], requiredFaces: p.projection === 'commands' ? requiredFacesOf(pack.faces.map((f) => f.id), pack) : [], projection: p.projection },
      limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: FRAGMENT, tilePx: 512 },
      notices: [...forensicNotices(p), ...(p.relaxed ? ['relaxed-session'] : [])],
      // Stage 4a: tiles mode only; a units-mode manifest is Stage 3's, key for key.
      ...(sd.sheetMode === 'tiles' ? {
        sheetMode: 'tiles' as const,
        sheets: (sd.doc.sheets ?? []).map((name, i) => ({ sheetId: String(i), name })),
        tiles: { px: TILE_PX, zoomSteps: [...TILE_ZOOM_STEPS], dprs: [...TILE_DPRS], palettes: [...TILE_PALETTES], maxBatch: TILE_BATCH_MAX },
        // Stage 4b (P1): only when the client asked for `sheet-view@1` (a 4a-opted session gets 4a's manifest, key for key).
        ...(sd.sheetView ? { sheetView: 1 as const } : {}),
      } : {}),
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
    const first = await obtainUnit(s.sid, sd, ui, sc, p.projection, missing, ws.data.gone.signal)
    if (!alive()) return
    if (typeof first === 'string') return refuse(ws, reqId, first)
    let unit: RenderedUnit = first
    let messages = unit.kind === 'commands' ? prepareUnit(reqId, ui, unit) : undefined
    if (!messages && unit.kind === 'commands') {
      // Assets too large for ASSET_DEFINE's caps: ship the unit as a raster (Task 7).
      const raster = await obtainUnit(s.sid, sd, ui, sc, 'final-raster', missing, ws.data.gone.signal)
      if (!alive()) return
      if (typeof raster === 'string') return refuse(ws, reqId, raster)
      unit = raster
    }
    let unmarked: string | undefined
    if (unit.kind === 'raster') {
      const d = await deliverRaster(alive, s.sid, sd, ui, sc, unit, (u) => prepareUnit(reqId, ui, u), ws.data.gone.signal)
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
  async function obtainUnit(sid: string, sd: { doc: DocState; policy: EffectivePolicy; decision: AuthDecision }, unitIndex: number, scale: number, projection: 'final-raster' | 'commands', missing?: string[], signal?: AbortSignal): Promise<RenderedUnit | RefusalCode> {
    const p = sd.policy, mark = markOf(p, sd.decision)
    const key = renderKey(keyFields(sd.doc, mark, unitIndex, scale, projection, missing))
    const hit = renderCache.get(key) as RenderedUnit | undefined
    if (hit) return hit
    const size = sd.doc.units[unitIndex]!, estimate = size.w * size.h * 4 * scale ** 2 * 2
    try {
      const unit = await o.pool.render(sid, { docRef: sd.doc.docRef, contentPolicy: p.contentPolicy, unitIndex, scale, mark, projection, ...(projection === 'commands' && missing ? { missingFaces: missing } : {}) }, estimate, signal)
      renderCache.put(key, unit, unit.kind === 'raster' ? unit.png.length : unit.ops.length + unit.assets.reduce((n, a) => n + (a.type === 'bitmap' ? a.rgba.length : 256), 0))
      return unit
    } catch (e) { return e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused' }
  }

  /** The shared per-unit tail (unit and print): Layer-1 over the canonical
   * raster, frame it with `pack`, then the durable delivery record — in that
   * order, so a unit over its frame cap leaves no delivery row, and nothing
   * goes out unrecorded. `undefined`: the caller died meanwhile (send nothing). */
  async function deliverRaster(alive: () => boolean, sid: string, sd: { doc: DocState; forensic?: Issued }, unitIndex: number, scale: number, unit: RenderedUnit & { kind: 'raster' }, pack: (u: RenderedUnit & { kind: 'raster' }) => Uint8Array[][] | undefined, signal?: AbortSignal): Promise<{ messages: Uint8Array[][]; unmarked?: string } | RefusalCode | undefined> {
    // Every raster this session receives — fallbacks included — carries its Layer-1 mark.
    let unmarked: string | undefined
    const canonical = sd.forensic ? unit.png : undefined
    if (sd.forensic) {
      const m = await markUnit(sid, sd, `${unitIndex}@${scale}`, unit, signal)
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
  /** `cancelled`: checked after the token await; true ⇒ nothing is sent (a print cancelled while it was validated).
   * `deny`: how a refusal is answered (default `refused`; notes answer `error`). */
  async function textOp(ws: WS, msg: Record<string, unknown>, cancelled?: () => boolean, deny: (reqId: unknown, code: 'policy-refused' | 'session-rotated') => void = (reqId, code) => refuse(ws, reqId, code)) {
    const reqId = msg.reqId
    const sid = ws.data.sid, s = sid ? table.get(sid) : undefined, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !s || !sd || !isU32(reqId) || reqId === 0) return void deny(reqId, 'policy-refused')
    const v = await verifyToken(o.keys.kToken, msg.token, { now: nowSec(), docRef: sd.doc.docRef, sid: s.sid })
    if (!owns(ws, s.sid) || cancelled?.()) return
    if (!v.ok) return void (v.reason === 'exp' ? close(ws, 'session-expired') : deny(reqId, 'policy-refused'))
    const jti = table.acceptJti(s.sid, v.claims.jti)
    if (jti === 'rotated') return void deny(reqId, 'session-rotated')
    if (jti === 'unknown') return void deny(reqId, 'policy-refused')
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
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job }, undefined, ws.data.gone.signal)
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
      job: (job) => o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job }, undefined, ws.data.gone.signal),
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
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job }, undefined, ws.data.gone.signal)
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
        return o.pool.text(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, ...job }, undefined, ws.data.gone.signal)
      },
      record: (r) => audit.recordCopy({ sid: s.sid, kind: 'a11y', paragraphs: r.paragraphs, fromUnit: args.unit, toUnit: args.unit }),
      alive: textAlive(ws, s.sid, sd),
      send: (m) => send(ws, m),
      refuse: (code, nums) => refuse(ws, reqId, code, nums),
    })
  }

  /**
   * Stage 4a (S4 §6.4, R14): `sheet-layout {reqId, token, sheetId}` on an XLSX tiles session (units mode refuses it
   * `unsupported-op` before this). In order: the token (`textOp`); `sheetId` the canonical decimal index of a visible
   * sheet, else `policy-refused`; the `'unit'` rate, charged once per request (its later parts charge nothing; S4
   * §6.8); the fingerprint and the document ledger (as `onUnit`); the build on the docRef's worker (F23); the encode
   * (an R14 cap ⇒ `layout-too-large`); then every part, while the session lives.
   */
  async function onSheetLayout(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    const sheets = sd.doc.sheets ?? []
    const id = msg.sheetId
    if (typeof id !== 'string' || !/^(0|[1-9][0-9]{0,8})$/.test(id) || Number(id) >= sheets.length) return refuse(ws, reqId, 'policy-refused')
    const rate = ws.data.rates.check('unit', now())
    if (rate === 'close') return close(ws, 'rate-limited')
    if (rate === 'refuse') return refuse(ws, reqId, 'rate-limited')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    const alive = textAlive(ws, s.sid, sd)
    let parts: object[]
    try {
      const base = await o.pool.sheetLayout(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, sheetIndex: Number(id), clock: sd.clock }, undefined, ws.data.gone.signal) as ServerSheetLayout
      if (!alive()) return
      const key = extentKey(sd, Number(id))
      if (!baseLayouts.has(key)) putBaseLayout(key, Promise.resolve(base))
      // Stage 4b (P11, F2): a sheet-view session's layout carries its current viewGen and (non-empty) session sizes; any
      // other session gets the identity layout, key for key 4a's. Read synchronously, so they agree with each other.
      const sizes = sd.sheetView ? sd.views.canonical(Number(id)) : []
      const layout = sessionLayout(base, sd.sheetView ? sd.views.viewGen(Number(id)) : 0, sizes)
      rememberExtent(sizedExtentKey(key, sizes), layout)
      parts = encodeSheetLayout(layout)
    } catch (e) {
      if (!alive()) return
      if (e instanceof LayoutTooLarge || (e instanceof RenderFailed && e.refusal === 'layout-too-large')) return refuse(ws, reqId, 'layout-too-large')
      return refuse(ws, reqId, e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
    }
    for (const p of parts) send(ws, { reqId, ...p })
  }

  /** Stage 4a (F8): the identity layout's cache key (the layout is the servable snapshot's, keyed like it by the session clock). */
  const extentKey = (sd: { doc: DocState; clock: SheetClock }, sheetIndex: number) => `${sd.doc.docRef}|${sd.doc.contentPolicyId}|${sd.doc.fingerprint}|${sheetIndex}|${sd.clock.day}|${sd.clock.timeZone}`
  /** Stage 4b: the extent cache key of a view: the base key, plus the hash of its canonical sizes when it has any (the
   * identity view keeps 4a's key). Never the viewGen: two views with equal sizes have one extent. */
  const sizedExtentKey = (base: string, sizes: WireSizes) => sizes.length ? `${base}|${createHash('sha256').update(JSON.stringify(sizes.map((e) => [e.axis, e.index, e.px]))).digest('hex')}` : base
  function extentOf(layout: StreamSheetLayout): SheetExtent {
    const m = streamSheetGeometry(layout).metrics
    return { colsCss: m.cols.totalCss(), rowsCss: m.rows.totalCss() }
  }
  function putExtent(key: string, p: Promise<SheetExtent>) {
    extents.delete(key); extents.set(key, p)
    while (extents.size > EXTENTS_MAX) extents.delete(extents.keys().next().value!)
    p.catch(() => { if (extents.get(key) === p) extents.delete(key) })
  }
  /** A `sheet-layout` answer gives the sheet's extent for free (the client asks for it before any tile). */
  function rememberExtent(key: string, layout: StreamSheetLayout) { if (!extents.has(key)) putExtent(key, Promise.resolve(extentOf(layout))) }
  function putBaseLayout(key: string, p: Promise<ServerSheetLayout>) {
    baseLayouts.delete(key); baseLayouts.set(key, p)
    while (baseLayouts.size > BASE_LAYOUTS_MAX) baseLayouts.delete(baseLayouts.keys().next().value!)
    p.catch(() => { if (baseLayouts.get(key) === p) baseLayouts.delete(key) })
  }
  /** Stage 4b: the sheet's identity layout from the docRef's worker, once per key (no socket's abort signal: other
   * sids may await the same promise). Rejects when the layout cannot be built. */
  function baseLayout(sid: string, sd: { doc: DocState; policy: EffectivePolicy; clock: SheetClock }, sheetIndex: number): Promise<ServerSheetLayout> {
    const key = extentKey(sd, sheetIndex)
    let p = baseLayouts.get(key)
    if (!p) putBaseLayout(key, p = o.pool.sheetLayout(sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, sheetIndex, clock: sd.clock }) as Promise<ServerSheetLayout>)
    return p
  }
  /** The sheet's whole domain in CSS px under the view's `sizes`, from its identity layout with the sizes merged (the
   * same axes `sheet-layout` gives the client, so the tile grid agrees), once per key. Rejects when the layout cannot
   * be built. */
  function sheetExtent(sid: string, sd: { doc: DocState; policy: EffectivePolicy; clock: SheetClock }, sheetIndex: number, sizes: WireSizes): Promise<SheetExtent> {
    const key = sizedExtentKey(extentKey(sd, sheetIndex), sizes)
    let p = extents.get(key)
    if (!p) putExtent(key, p = baseLayout(sid, sd, sheetIndex).then((l) => extentOf(sessionLayout(l, 0, sizes))))
    return p
  }

  /**
   * Stage 4a (S4 §6.5, R4): `tiles {reqId, token, sheetId, viewGen, zoomStep, dpr, palette, tiles}` on an XLSX tiles
   * session. The batch is registered on arrival, before any await, so a `sheet-cancel` right behind it finds it.
   */
  async function onTiles(ws: WS, msg: Record<string, unknown>) {
    const sid0 = ws.data.sid, sd0 = sid0 ? sessionDoc.get(sid0) : undefined
    const reqId = msg.reqId
    if (ws.data.phase !== 'stream' || !sid0 || !sd0 || !owns(ws, sid0) || !isU32(reqId) || reqId === 0) return refuse(ws, reqId, 'policy-refused')
    // A reqId already in progress on the sid would make its replies ambiguous.
    if (sd0.tiles.batches.has(reqId)) return refuse(ws, reqId, 'policy-refused')
    const batch: TileBatch = { reqId, stopped: false }
    sd0.tiles.batches.set(reqId, batch)
    try { await tilesOp(ws, msg, sd0, batch) } finally { if (sd0.tiles.batches.get(reqId) === batch) sd0.tiles.batches.delete(reqId) }
  }
  /**
   * In order, each failure answering one whole-request `refused` that charges nothing: the shape (`tilesArgs`:
   * `policy-refused`, or `stale-view` for another `viewGen`); the token (`textOp`); the fingerprint (else close
   * `layout-changed`) and the document ledger; every coordinate on the sheet's grid (F8: `tileLimit` of its extent).
   * Then the `'tile'` rate per tile (R4: the rest are refused `rate-limited` in `tiles-done`, with `retryAfterMs`); the
   * `sheet_views` row, once per `(sid, sheetId, viewGen)`; each admitted tile (`runTiles`, at most
   * `perSessionOutstanding` in flight per sid): the canonical tile (cache, else `pool.tile`), Layer-1 on its own
   * basis (`tileBasisLabel`), the frame, the `tile_deliveries` row, the send; then `tiles-done`. A batch stopped
   * before the rate charge (a `sheet-cancel` during validation) charges and sends nothing but `tiles-done`.
   */
  async function tilesOp(ws: WS, msg: Record<string, unknown>, sd0: { doc: DocState; views: SessionViews }, batch: TileBatch) {
    const reqId = batch.reqId
    const args = tilesArgs(msg, (sd0.doc.sheets ?? []).length, (i) => sd0.views.viewGen(i))
    if (typeof args === 'string') return refuse(ws, reqId, args)
    // Stage 4b: the view's sizes, read in the same tick as its viewGen; the one value this batch's extent, keys, jobs and
    // `sheet_views` row use. A later resize of the sheet stops the batch (P14).
    const sizes = sd0.views.canonical(args.sheetIndex)
    batch.sheetIndex = args.sheetIndex
    const a = await textOp(ws, msg)
    if (!a) return
    const { s, sd } = a
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
    const alive = () => textAlive(ws, s.sid, sd)()
    let extent: SheetExtent
    // A sheet whose layout cannot be built (over an R14 cap, a diagnostic geometry) has no grid to bound tiles on.
    try { extent = await sheetExtent(s.sid, sd, args.sheetIndex, sizes) } catch (e) { return alive() ? refuse(ws, reqId, e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined }
    if (!alive()) return
    if (!tilesInGrid(args, extent)) return refuse(ws, reqId, 'policy-refused')
    const done = (refused: Array<TileCoord & { code: TileRefusalCode }>, retryAfterMs: number) => {
      if (alive()) send(ws, { reqId, type: 'tiles-done', refused, ...(retryAfterMs > 0 ? { retryAfterMs } : {}), ...(batch.stopped ? { cancelled: true } : {}) })
    }
    if (batch.stopped) return done([], 0)
    const { admitted, retryAfterMs } = ws.data.rates.takeTiles(args.tiles.length, now())
    const timing = o.onTileTiming, t0 = timing ? performance.now() : 0
    const run = args.tiles.slice(0, admitted)
    const limited = args.tiles.slice(admitted).map((t) => ({ col: t.col, row: t.row, code: 'rate-limited' as const }))
    // R7: the view and clock this sid's tiles of the sheet are drawn under, durable before its first tile.
    const viewKey = `${args.sheetId}|${args.viewGen}`
    if (run.length && !sd.tiles.views.has(viewKey)) {
      try { audit.recordSheetView({ sid: s.sid, sheetId: args.sheetId, viewGen: args.viewGen, spec: { sort: null, filters: [], sizes }, clock: { day: sd.clock.day, timeZone: sd.clock.timeZone } }) } catch { return refuse(ws, reqId, 'policy-refused') }
      sd.tiles.views.add(viewKey)
    }
    const refused = await runTiles({
      batch, tiles: run, slots: sd.tiles.slots, alive,
      deliver: (t) => deliverTile(alive, s.sid, sd, args, sizes, t, reqId, ws.data.gone.signal),
      send: (frames) => { sendFrames(ws, frames); if (timing) timing({ t0, t1: performance.now(), zoomStep: args.zoomStep, dpr: args.dpr }) },
    })
    done([...refused, ...limited], limited.length ? retryAfterMs : 0)
  }

  /** The canonical-tier key of one tile (S4 §6.5): the §9.3 fields with `policyDetails.viewStateHash` (the view's sizes,
   * Stage 4b, and the session clock), `scaleBucket` = dpr, `unitId` = sheetIndex, and the `tile` field. Never the sid
   * and never the viewGen: sessions with equal sizes share the canonical tile, different sizes never do. */
  function tileKeyFields(doc: DocState, mark: MarkSpec, clock: SheetClock, a: TilesArgs, sizes: WireSizes, t: TileCoord): RenderKeyFields {
    const unit = keyFields(doc, mark, a.sheetIndex, a.dpr, 'final-raster')
    return { ...unit, policyDetails: { projection: doc.contentPolicyId, viewStateHash: viewStateHash({ sort: null, filters: [], sizes, authorHiddenTier: 'none', clock }) }, tile: { sheetIndex: a.sheetIndex, zoomStep: a.zoomStep, dpr: a.dpr, palette: a.palette, col: t.col, row: t.row } }
  }
  /** One tile for a session: the canonical tile (the render cache, else the pool, then cached), Layer-1 over it on
   * the tile's own basis, the frame, then the durable `tile_deliveries` row (forensic sessions, as units). In that
   * order, so a tile over its frame cap leaves no row. `undefined`: the caller died meanwhile (send nothing). */
  async function deliverTile(alive: () => boolean, sid: string, sd: { doc: DocState; policy: EffectivePolicy; decision: AuthDecision; clock: SheetClock; forensic?: Issued }, a: TilesArgs, sizes: WireSizes, t: TileCoord, reqId: number, signal: AbortSignal): Promise<Uint8Array[][] | TileRefusalCode | undefined> {
    // A long batch can outlive the parse: every tile re-charges it.
    if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return 'policy-refused'
    const mark = markOf(sd.policy, sd.decision)
    const key = renderKey(tileKeyFields(sd.doc, mark, sd.clock, a, sizes, t))
    let tile = renderCache.get(key) as RenderedTile | undefined
    if (!tile) {
      try { tile = await o.pool.tile(sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, sheetIndex: a.sheetIndex, viewGen: a.viewGen, zoomStep: a.zoomStep, dpr: a.dpr, palette: a.palette, col: t.col, row: t.row, mark, clock: sd.clock, sizes }, TILE_BYTES_ESTIMATE, signal) } catch (e) { return !alive() ? undefined : e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused' }
      // Fail closed: only the tile asked for (F24: its palette too) is cached or sent.
      if (tile.w !== TILE_PX || tile.h !== TILE_PX || tile.palette !== a.palette) return 'policy-refused'
      renderCache.put(key, tile, tile.png.length)
    }
    if (!alive()) return undefined
    const canonical = tile.png
    let png = canonical, unmarked: string | undefined
    if (sd.forensic) {
      const m = await markUnit(sid, sd, tileBasisLabel(a.sheetId, a.zoomStep, a.dpr, t.col, t.row), { png: canonical, w: tile.w, h: tile.h }, signal)
      if (!alive()) return undefined
      if (typeof m === 'string') return m === 'rate-limited' ? 'rate-limited' : 'policy-refused'
      png = m.unit.png; unmarked = m.unmarked
    }
    let body: Uint8Array
    try { body = encodeTileBody({ sheetIndex: a.sheetIndex, viewGen: a.viewGen, zoomStep: a.zoomStep, dpr: a.dpr, col: t.col, row: t.row, png, w: tile.w, h: tile.h }) } catch { return 'policy-refused' }
    const f = frame(KIND.XLSX_TILE, reqId, body)
    if (!f) return 'policy-refused'
    // Durable before the tile goes out (never a delivered mark without a record); a failed write refuses it.
    if (sd.forensic) {
      try { audit.recordTileDelivery({ sid, sheetId: a.sheetId, viewGen: a.viewGen, zoomStep: a.zoomStep, dpr: a.dpr, palette: a.palette, col: t.col, row: t.row, status: unmarked ? 'unmarked' : 'l1', ...(unmarked ? { reason: unmarked } : {}), refSha256: createHash('sha256').update(canonical).digest('hex') }) } catch { return 'policy-refused' }
    }
    return [f]
  }

  /** Stage 4a (R12): `sheet-cancel {reqId, target}`: no tile of batch `target` starts after this; its `tiles-done`
   * carries `cancelled: true`. An unknown target is ignored, and nothing answers the cancel itself. */
  function onSheetCancel(ws: WS, msg: Record<string, unknown>) {
    const sid = ws.data.sid, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !sid || !sd || !owns(ws, sid) || !isU32(msg.target)) return
    const b = sd.tiles.batches.get(msg.target)
    if (b) b.stopped = true
  }

  /**
   * Stage 4b (P11–P14): `sheet-resize {reqId, token, sheetId, viewGen, axis, index, size}` or `{…, reset: true}` on a
   * `sheet-view@1` session. In order, each failure one `refused` that changes nothing:
   * - the token (`textOp`);
   * - the arguments (`resizeArgs`), else `policy-refused`;
   * - the `'sheet'` rate (refuse-only), then the one `'sheet'` op in flight (`sheetBusy`), both `rate-limited`; the flag
   *   is set here, before the op's next await, and its `finally` clears it only while this socket still owns the sid
   *   (a superseded socket's late op never drops the new socket's flag);
   * - `stale-view` for a `viewGen` other than the sheet's;
   * - the fingerprint (else close `layout-changed`) and the document ledger, then the sheet's identity layout;
   * - `hidden-target` when the row or column's authored size is 0 (a collapsed run of the layout);
   * - `coverage-unknown` under `'visible-only'` when the sheet's coverage is unknown;
   * - `views.apply`, which may refuse `view-state-limit`.
   * Then a change stops the sid's tile batches of that sheet (P14: their late tiles carry the old viewGen) and the
   * answer is `sheet-view {sheetId, viewGen, change}` with the applied (clamped) size. Nothing is applied once the
   * socket no longer owns the sid.
   */
  async function onSheetResize(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    const args = resizeArgs(msg, (sd.doc.sheets ?? []).length)
    if (!args) return refuse(ws, reqId, 'policy-refused')
    const rates = ws.data.rates
    if (rates.check('sheet', now()) !== 'ok' || rates.sheetBusy) return refuse(ws, reqId, 'rate-limited')
    rates.sheetBusy = true
    try {
      if (args.viewGen !== sd.views.viewGen(args.sheetIndex)) return refuse(ws, reqId, 'stale-view')
      if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
      if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
      const alive = textAlive(ws, s.sid, sd)
      let base: ServerSheetLayout
      try { base = await baseLayout(s.sid, sd, args.sheetIndex) } catch (e) { return alive() ? refuse(ws, reqId, e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined }
      if (!alive()) return
      if (!('reset' in args)) {
        const runs = args.axis === 'row' ? base.collapsedRows : base.collapsedCols
        if (runs.some((r) => r.start <= args.index && args.index <= r.end)) return refuse(ws, reqId, 'hidden-target')
      }
      if (sd.policy.contentPolicy === 'visible-only' && base.coverageUnknown) return refuse(ws, reqId, 'coverage-unknown')
      // Synchronous from the viewGen re-check to the send: no other op can interleave.
      if (args.viewGen !== sd.views.viewGen(args.sheetIndex)) return refuse(ws, reqId, 'stale-view')
      const r = sd.views.apply(args.sheetIndex, 'reset' in args ? { reset: true } : { axis: args.axis, index: args.index, size: args.size })
      if (r === 'view-state-limit') return refuse(ws, reqId, 'view-state-limit')
      if (r.viewGen !== args.viewGen) for (const b of sd.tiles.batches.values()) if (b.sheetIndex === args.sheetIndex) b.stopped = true
      send(ws, { reqId, type: 'sheet-view', sheetId: args.sheetId, viewGen: r.viewGen, change: 'reset' in args ? { reset: true } : { axis: args.axis, index: args.index, size: r.applied! } })
    } finally {
      if (owns(ws, s.sid)) rates.sheetBusy = false
    }
  }

  /**
   * Stage 4b (P9, P12): `sheet-fit {reqId, token, sheetId, viewGen, axis, index, pane, regions?}` on a `sheet-view@1`
   * session. In `onSheetResize`'s order, each failure one `refused` that changes nothing: the token; the arguments
   * (`fitArgs`), else `policy-refused`; the `'sheet'` rate, then `sheetBusy` (`rate-limited`; set before the next await,
   * cleared in the `finally` only while this socket owns the sid); `stale-view`; the fingerprint and the document ledger;
   * `hidden-target` for a collapsed row or column of the identity layout. Then the `sheet-text` fit job over the sheet's
   * text view at the session's sizes (read in the same tick as the viewGen check), answered `sheet-fit-result {size}`.
   * No `coverage-unknown`: an unknown-coverage sheet's text view is blank, so it fits to the blank size. Fit never
   * changes view state; it is served with copy and a11y off.
   */
  async function onSheetFit(ws: WS, msg: Record<string, unknown>) {
    const a = await textOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    const args = fitArgs(msg, (sd.doc.sheets ?? []).length)
    if (!args) return refuse(ws, reqId, 'policy-refused')
    const rates = ws.data.rates
    if (rates.check('sheet', now()) !== 'ok' || rates.sheetBusy) return refuse(ws, reqId, 'rate-limited')
    rates.sheetBusy = true
    try {
      if (args.viewGen !== sd.views.viewGen(args.sheetIndex)) return refuse(ws, reqId, 'stale-view')
      // The sizes of that viewGen: no resize can land meanwhile (one 'sheet' op in flight).
      const sizes = sd.views.canonical(args.sheetIndex)
      if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
      if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return refuse(ws, reqId, 'policy-refused')
      const alive = textAlive(ws, s.sid, sd)
      let base: ServerSheetLayout
      try { base = await baseLayout(s.sid, sd, args.sheetIndex) } catch (e) { return alive() ? refuse(ws, reqId, e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined }
      if (!alive()) return
      const runs = args.axis === 'row' ? base.collapsedRows : base.collapsedCols
      if (runs.some((r) => r.start <= args.index && args.index <= r.end)) return refuse(ws, reqId, 'hidden-target')
      let r: SheetTextReply
      try {
        r = await o.pool.sheetText(s.sid, { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, clock: sd.clock, op: 'fit', sheetIndex: args.sheetIndex, axis: args.axis, index: args.index, pane: args.pane, ...(args.regions?.length ? { regions: args.regions } : {}), sizes }, undefined, ws.data.gone.signal)
      } catch (e) { return alive() ? refuse(ws, reqId, e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined }
      if (!alive()) return
      // Fail closed: only a fit answer with a finite size travels.
      if (r.op !== 'fit' || !Number.isFinite(r.size) || r.size < 0) return refuse(ws, reqId, 'policy-refused')
      send(ws, { reqId, type: 'sheet-fit-result', size: r.size })
    } finally {
      if (owns(ws, s.sid)) rates.sheetBusy = false
    }
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

  /** §12a.5/§12a.8 notes (Stage 3c): a presenter's notes image of one unit.
   * Checks in the plan's order, all before any render: token → `canNotes`
   * (and a deck) → unit in range and not hidden → scale → the 'notes' rate.
   * Then the plan (size, characters drawn), the durable `notes_deliveries`
   * row, the draw with the visible mark, Layer-1, the frame — and only then
   * `notes-start` and the kind-7 frame. A refusal is one `error` message: no
   * frame; before the row, no row. Never cached across requests (C4). */
  async function onNotes(ws: WS, msg: Record<string, unknown>) {
    const deny = (reqId: unknown, code: NotesErrorCode) => send(ws, { reqId: isU32(reqId) ? reqId : 0, type: 'error', code })
    const a = await textOp(ws, msg, undefined, deny)
    if (!a) return
    const { reqId, s, sd } = a
    if (!canNotes(sd.policy) || sd.doc.format !== 'pptx') return deny(reqId, 'policy-refused')
    const unit = msg.unit, scale = msg.scale
    if (!Number.isInteger(unit) || (unit as number) < 0 || (unit as number) >= sd.doc.units.length || sd.doc.hidden.has(unit as number)) return deny(reqId, 'invalid-unit')
    if (!(SCALE_BUCKETS as readonly unknown[]).includes(scale)) return deny(reqId, 'policy-refused')
    // A 'notes' overflow is a refusal only (never a breach toward 4008).
    if (ws.data.rates.check('notes', now()) !== 'ok') return deny(reqId, 'rate-limited')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    const alive = textAlive(ws, s.sid, sd)
    const ui = unit as number, sc = scale as 1 | 2
    const job = { docRef: sd.doc.docRef, contentPolicy: sd.policy.contentPolicy, unitIndex: ui, scale: sc, mark: markOf(sd.policy, sd.decision) }
    const run = async <T extends NotesPlan | NotesRaster>(draw: boolean, estimate: number): Promise<T | 'rate-limited' | 'policy-refused' | undefined> => {
      if (!ledger.touch(sd.doc.docRef, sd.policy.contentPolicy, sd.doc.bytes.length)) return 'policy-refused'
      let r: NotesPlan | NotesRaster
      try { r = await o.pool.notes(s.sid, { ...job, draw }, estimate, ws.data.gone.signal) } catch (e) { return alive() ? (e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined }
      return alive() ? r as T : undefined
    }
    const plan = await run<NotesPlan>(false, NOTES_PLAN_BYTES)
    if (plan === undefined) return
    if (typeof plan === 'string') return deny(reqId, plan)
    // Durable before the draw and before anything about it goes out; a failed write refuses it.
    try { audit.recordNotes({ sid: s.sid, unitIndex: ui, scale: sc, charsDrawn: plan.chars }) } catch { return deny(reqId, 'policy-refused') }
    const drawn = await run<NotesRaster>(true, plan.w * plan.h * 4 * 2)
    if (drawn === undefined) return
    if (typeof drawn === 'string') return deny(reqId, drawn)
    // Fail closed: the image must be the one planned and audited.
    if (drawn.kind !== 'raster' || drawn.w !== plan.w || drawn.h !== plan.h || drawn.chars !== plan.chars) return deny(reqId, 'policy-refused')
    let img: NotesRaster = drawn
    if (sd.forensic) {
      const m = await markUnit(s.sid, sd, notesBasisLabel(ui, sc), drawn, ws.data.gone.signal)
      if (!alive()) return
      if (typeof m === 'string') return deny(reqId, m === 'rate-limited' ? m : 'policy-refused')
      // Both marks on every notes image (§13.1): one that cannot carry Layer-1 is not sent.
      if (m.unmarked) return deny(reqId, 'policy-refused')
      img = m.unit
    }
    const f = frame(KIND.NOTES_RASTER, reqId, encodeNotesBody(ui, img.w, img.h, img.png))
    if (!f) return deny(reqId, 'policy-refused')
    // The marked delivery trace attributes (its notes basis and canonical): durable before the send, after the frame
    // (an image over its cap leaves no delivery row); a failed write refuses it.
    if (sd.forensic) {
      try { audit.recordNotesMark({ sid: s.sid, unitIndex: ui, scale: sc, refSha256: createHash('sha256').update(drawn.png).digest('hex') }) } catch { return deny(reqId, 'policy-refused') }
    }
    send(ws, { reqId, type: 'notes-start', w: img.w, h: img.h, scale: sc })
    sendFrames(ws, [f])
  }

  // -------------------------------------------------------------------------
  // Piece 4 Task 7: show-plan, show-track, show-morph (plan D1, D4, D7, D11)
  // -------------------------------------------------------------------------

  /** One asset on its way out: its kind-6 tag, ids and pixels, and its Layer-1 crop (`rest`, page device px). */
  interface OutAsset { tag: 2 | 3 | 4; assetId: number; w: number; h: number; png: Uint8Array; rest: { x: number; y: number; scale: number } }
  // The admitted bytes are also the build's preparation grant on the worker (PR #78 round 13, `showBuildGrant`).
  const showEstimate = (size: UnitSize, scale: number) => showBuildGrant(size, scale)
  const shownUnit = (d: DocState, u: unknown): u is number => Number.isInteger(u) && (u as number) >= 0 && (u as number) < d.units.length && !d.hidden.has(u as number)
  const showFailure = (e: unknown): ShowErrorCode => e instanceof AdmissionRefused ? 'rate-limited' : e instanceof RenderTimeout ? 'show-plan-timeout' : e instanceof RenderFailed && e.refusal === 'unit-range' ? 'invalid-unit' : 'policy-refused'

  /**
   * The three show ops. Their rules run in this order, all before any render (Task 7 brief):
   * 1. the token, then `can.present` (a deck);
   * 2. the unit: in range, not hidden, animated (D5); a track's step and a Morph's pair (`to` entered by Morph,
   *    `from ≠ to`, both shown, `fromStep` an integer in 0..64; its exact range against the from-slide's steps is
   *    checked on the worker, after the charge);
   * 3. `scale` in `SCALE_BUCKETS` (1 or 2) and `motion` valid (a Morph under `reduced` motion plays as a fade, D13:
 *    `invalid-unit`);
   * 4. the `'unit'` rate: one charge per op (owner decision 2), `close` → 4008, `refuse` → `rate-limited`;
   * 5. the fingerprint and the document ledger (as `onUnit`);
   * 6. the build on the docRef's worker (L1/L2 caches, D6);
   * 7. Layer-1 per asset with the session's basis (`markLayer`, D7); an asset that keeps no carrier ships unmarked;
   * 8. every message framed, then the durable `show_deliveries` row;
   * 9. `diag` text (forensic-unmarked, Morph diagnostics), `*-start`, then the kind-6 messages, each one sent only
   *    while the session lives.
   */
  /** Task 11b fix round: the session's record of what the plan bound for `unit` ships, and how (`MediaSlot.plans`). */
  const planned = (sd: { media: MediaSlot }, unit: number, plan: ShowPlanRecord) => {
    sd.media.plans.set(unit, new Map((plan.media ?? []).map((r) => [r.mediaId, r.refused])))
    // Encode-ahead of this unit waited for its plan (Important 4): the queue looks again.
    sd.media.queue?.wake()
  }

  async function onShow(ws: WS, msg: Record<string, unknown>, kind: ShowDeliveryRecord['kind']) {
    const deny = (reqId: unknown, code: ShowErrorCode) => send(ws, { reqId: isU32(reqId) ? reqId : 0, type: 'error', code })
    const a = await textOp(ws, msg, undefined, deny)
    if (!a) return
    const { reqId, s, sd } = a
    const d = sd.doc
    if (d.format !== 'pptx') return deny(reqId, 'policy-refused')
    // D19: a unit with `anim`, or one the manifest marked `media` (only with `can.media`).
    const animated = (u: unknown): u is number => shownUnit(d, u) && (d.anim.has(u) || (canMedia && d.media.has(u)))
    const step = (x: unknown, max: number): x is number => Number.isInteger(x) && (x as number) >= 0 && (x as number) <= max
    let unit: number
    if (kind === 'morph') {
      if (!shownUnit(d, msg.from) || !animated(msg.to) || !d.morph.has(msg.to) || msg.from === msg.to || !step(msg.fromStep, SHOW_PLAN_LIMITS_STEPS)) return deny(reqId, 'invalid-unit')
      unit = msg.to
    } else {
      if (!animated(msg.unit)) return deny(reqId, 'invalid-unit')
      if (kind === 'track' && !step(msg.step, SHOW_PLAN_LIMITS_STEPS - 1)) return deny(reqId, 'invalid-unit')
      unit = msg.unit
    }
    const scale = msg.scale, motion = msg.motion
    if (!(SCALE_BUCKETS as readonly unknown[]).includes(scale) || (scale !== 1 && scale !== 2) || (motion !== 'full' && motion !== 'reduced')) return deny(reqId, 'policy-refused')
    // D13: under reduced motion a Morph plays as a short fade (`reducedTransition`, as local mode): no pair plays.
    if (kind === 'morph' && motion === 'reduced') return deny(reqId, 'invalid-unit')
    const rate = ws.data.rates.check('unit', now())
    if (rate === 'close') return close(ws, 'rate-limited')
    if (rate === 'refuse') return deny(reqId, 'rate-limited')
    if (fingerprintOf(d.contentPolicyId) !== d.fingerprint) return close(ws, 'layout-changed')
    if (!ledger.touch(d.docRef, sd.policy.contentPolicy, d.bytes.length)) return deny(reqId, 'policy-refused')
    const alive = textAlive(ws, s.sid, sd)
    const p = sd.policy, mark = markOf(p, sd.decision), strictMarks = o.animation?.strictMarks === true
    const base = { docRef: d.docRef, contentPolicy: p.contentPolicy, scale: scale as 1 | 2, motion: motion as 'full' | 'reduced', mark, strictMarks, media: canMedia }
    // A Morph pair prepares both slides (PR #78 round 13): it is admitted for both.
    const estimate = kind === 'morph' ? showMorphGrant(d.units[msg.from as number]!, d.units[unit]!, scale) : showEstimate(d.units[unit]!, scale)

    // 6. The build.
    // `head`: the plan (tag 1) or Morph record (tag 5); a track has none.
    let head: ShowMessage | undefined, assets: OutAsset[], page: { w: number; h: number }, decoded: number
    const diags: Array<Record<string, unknown>> = []
    const start: Record<string, unknown> = { reqId, type: `show-${kind}-start`, unit }
    try {
      if (kind === 'plan') {
        const built = await o.pool.showPlan(s.sid, { ...base, unitIndex: unit }, estimate, ws.data.gone.signal)
        if (!alive()) return
        if ('refused' in built) {
          // Task 11b: a still with audio answers its sound-only plan (one message, no asset): it plays its piece-3
          // still and sounds. Any other refusal, or a still without audio (or media off), is piece 4's error.
          if (!built.sounds) return deny(reqId, built.refused)
          head = { tag: 1, plan: bindMedia(built.sounds.plan, (ref) => sessionMediaId(sd.mediaKey, mediaIdentity(ref)), built.sounds.mediaRefs) }
          assets = []
          page = { w: 0, h: 0 }
          decoded = built.sounds.decodedBytes
        } else {
          // D2: the cached plan's placeholders become this session's media ids.
          head = { tag: 1, plan: bindMedia(built.plan, (ref) => sessionMediaId(sd.mediaKey, mediaIdentity(ref)), built.mediaRefs) }
          assets = built.assets
          const entry = built.assets.find((x) => x.tag === 3)
          if (!entry) return deny(reqId, 'policy-refused')
          page = { w: entry.w, h: entry.h }
          decoded = built.decodedBytes
        }
      } else if (kind === 'track') {
        const frames = await o.pool.showTrack(s.sid, { ...base, unitIndex: unit, step: msg.step as number }, estimate, ws.data.gone.signal)
        if (!alive()) return
        if (!frames.length) return deny(reqId, 'invalid-unit')
        assets = frames.map((f) => ({ tag: 4, ...f, rest: { x: 0, y: 0, scale: 1 } }))
        page = { w: frames[0]!.w, h: frames[0]!.h }
        decoded = trackWindowBytes(frames)
        start.step = msg.step
      } else {
        const built = await o.pool.showMorph(s.sid, { ...base, from: msg.from as number, to: unit, fromStep: msg.fromStep as number }, estimate, ws.data.gone.signal)
        if (!alive()) return
        if ('refused' in built) return deny(reqId, built.refused)
        head = { tag: 5, morph: built.record }
        assets = 'frames' in built ? built.frames.map((f) => ({ tag: 4 as const, ...f })) : built.assets.map((x) => ({ tag: 2 as const, ...x }))
        page = built.page
        // A composited pair plays through the D23 window like a track; a layered pair's assets are all held at once.
        decoded = 'frames' in built ? trackWindowBytes(assets) : assets.reduce((n, x) => n + x.w * x.h * 4, 0)
        for (const code of built.diagnostics) diags.push({ reqId, type: 'diag', code, unit, ...(code === 'morph-composited' && 'reason' in built ? { reason: built.reason } : {}) })
        Object.assign(start, { from: msg.from, to: unit, fromStep: msg.fromStep })
      }
    } catch (e) { return alive() ? deny(reqId, showFailure(e)) : undefined }

    // 7. Layer-1 per asset, in order (the scratch lease holds two at most).
    const marked = await markShow(alive, s.sid, sd, unit, scale, page, assets, ws.data.gone.signal)
    if (marked === undefined) return
    if (typeof marked === 'string') return deny(reqId, marked)

    // 8. Every message framed (none over its cap), then the durable row.
    const payloads: Uint8Array[] = []
    try {
      if (head) payloads.push(encodeShowMessage(head))
      for (const x of marked.assets) payloads.push(encodeShowMessage({ tag: x.tag, unit, assetId: x.assetId, w: x.w, h: x.h, png: x.png }))
    } catch (e) { if (e instanceof PlanRejected) return deny(reqId, 'policy-refused'); throw e }
    // The Morph record's bytes count toward the reply's decoded total (a plan's are in `decodedBytes` already).
    if (kind === 'morph') decoded += payloads[0]!.length
    const messages: Uint8Array[][] = []
    for (const pl of payloads) { const f = frame(KIND.SHOW_PLAN, reqId, pl); if (!f) return deny(reqId, 'policy-refused'); messages.push(f) }
    const unmarked = marked.unmarked.length
    try { audit.recordShow({ sid: s.sid, unitIndex: unit, kind, assetCount: assets.length, unmarked, ...(unmarked ? { reason: marked.reason } : {}) }) } catch { return deny(reqId, 'policy-refused') }
    // Fix round 1 (Important 3): the plan's media answers count from here on, once nothing can refuse the reply.
    if (head?.tag === 1) planned(sd, unit, head.plan)

    // 9. Diagnostics, the start, the messages.
    if (marked.reason === 'layer-capacity' && unmarked) diags.push({ reqId, type: 'diag', code: 'forensic-unmarked', unit, reason: 'layer-capacity', assetIds: marked.unmarked })
    for (const x of diags) send(ws, x)
    // D23: a track or Morph start also declares Σ its encoded payloads (retained while it plays); a plan's never.
    send(ws, { ...start, messages: messages.length, decodedBytes: decoded, ...(kind === 'plan' ? {} : { encodedBytes: payloads.reduce((n, p) => n + p.length, 0) }) })
    for (const m of messages) {
      if (!alive()) return
      if (!sendFrames(ws, [m])) return
    }
  }

  /**
   * D7: Layer-1 on each show asset with the session's page-keyed basis (`${unit}@${scale}`), cropped to the asset's
   * rest. An asset left with no carrier ships its visibly marked bytes, listed in `unmarked` (`layer-capacity`); a
   * session without Layer-1 (policy, or `testOnly.forensicOff`) lists every asset (`forensic-off`). `undefined`: the
   * session died meanwhile (send nothing).
   */
  async function markShow(alive: () => boolean, sid: string, sd: { doc: DocState; forensic?: Issued }, unit: number, scale: number, page: { w: number; h: number }, assets: OutAsset[], signal?: AbortSignal): Promise<{ assets: OutAsset[]; unmarked: number[]; reason: 'layer-capacity' | 'forensic-off' } | ShowErrorCode | undefined> {
    if (!sd.forensic || o.testOnly?.forensicOff === true) return { assets, unmarked: assets.map((x) => x.assetId), reason: 'forensic-off' }
    const kBasis = deriveBasisKey(o.keys.kBasisRoot, sd.doc.docRef, `${unit}@${scale}`)
    const out: OutAsset[] = [], unmarked: number[] = []
    for (const x of assets) {
      let release: () => void
      try { release = await leaseFor(sid).acquire() } catch (e) { return e instanceof LeaseFull ? 'rate-limited' : 'policy-refused' }
      let r
      try {
        r = await o.pool.markLayer(sid, { png: x.png, w: x.w, h: x.h, kBasis, code: sd.forensic.code, crop: { x: x.rest.x, y: x.rest.y, pageW: page.w, pageH: page.h, scale: x.rest.scale }, ...(o.encodeHangMs ? { hangMs: o.encodeHangMs } : {}) }, encodeEstimate(x.w, x.h, x.png.length), signal)
      } catch (e) { return alive() ? (e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused') : undefined } finally { release() }
      if (!alive()) return undefined
      if ('unmarked' in r) { forensicStats.unmarked++; unmarked.push(x.assetId); out.push(x) } else { forensicStats.marked++; out.push({ ...x, png: r.png }) }
    }
    return { assets: out, unmarked, reason: 'layer-capacity' }
  }

  // -------------------------------------------------------------------------
  // Piece 5 Task 9: media-show, media-cursor, media-request, media-end (spec §4.3; plan D1, D12, D13, D23)
  // -------------------------------------------------------------------------

  type MediaReason = MediaRefused['reason']
  /** `media-refused`; `mediaId` is echoed only when it is a well-formed id (a client validates every reply). */
  const mediaDeny = (ws: WS, reqId: unknown, reason: MediaReason, mediaId?: string) =>
    send(ws, { reqId: isU32(reqId) ? reqId : 0, type: 'media-refused', ...(mediaId !== undefined && MEDIA_ID.test(mediaId) ? { mediaId } : {}), reason })
  /** The token preamble of a media op: a stale jti answers the shared `session-rotated` refusal (the client's retry
   * signal, as every op); anything else is `media-refused policy-refused`. */
  const mediaOp = (ws: WS, msg: Record<string, unknown>) =>
    textOp(ws, msg, undefined, (reqId, code) => (code === 'session-rotated' ? refuse(ws, reqId, code) : mediaDeny(ws, reqId, code)))

  /**
   * `media-show {reqId, token, unit}` (D1): the token, `can.media` (and a deck), the unit (shown); then the session's
   * queue (created on the first show, kept until `media-end`) starts from `unit`. While it loads the slides' items it
   * records every audio use of the session's plans in `sd.media.uses` (refused ones included), the only ids
   * `media-request` serves. Counted against the message rate only. Answers `media-show-ok {reqId, queued}`.
   */
  async function onMediaShow(ws: WS, msg: Record<string, unknown>) {
    const a = await mediaOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    let unit: number
    try { unit = parseMediaShow(msg).unit } catch { return mediaDeny(ws, reqId, 'policy-refused') }
    const svc = services, info = o.media?.info
    if (!svc || !info || sd.doc.format !== 'pptx') return mediaDeny(ws, reqId, 'media-unavailable')
    if (!shownUnit(sd.doc, unit)) return mediaDeny(ws, reqId, 'policy-refused')
    if (fingerprintOf(sd.doc.contentPolicyId) !== sd.doc.fingerprint) return close(ws, 'layout-changed')
    let q = sd.media.queue
    if (!q) {
      const docRef = sd.doc.docRef, policy = sd.policy.contentPolicy, uses = sd.media.uses
      try {
        q = svc.createQueue({
          sid: s.sid, units: sd.doc.units.length,
          items: async (u) => {
            let refs = sd.media.refs.get(u)
            if (!refs) {
              // A long show can outlive the parse: every read re-charges it.
              if (!ledger.touch(docRef, policy, sd.doc.bytes.length)) throw new Error('document cache full')
              refs = await o.pool.mediaRefs(s.sid, docRef, policy, u)
              sd.media.refs.set(u, refs)
            }
            if (!refs.hidden) for (const r of refs.refs) {
              if (r.hiddenShape) continue
              const id = sessionMediaId(sd.mediaKey, r.identity)
              let byUnit = uses.get(id)
              if (!byUnit) uses.set(id, byUnit = new Map())
              byUnit.set(u, { unit: u, ...('refused' in r.entry ? { refused: r.entry.refused } : {}) })
            }
            const items = mediaJobsFor(refs, u, { mediaKey: sd.mediaKey, docId: docRef }, info)
            for (const it of items) { const use = uses.get(it.mediaId)?.get(u); if (use) use.item = it }
            return items
          },
          // Task 11b fix round: encode-ahead skips a use the plan sent for its unit refuses (a request is refused too).
          // Fix round 1 (Important 4): only a use the plan sent for its unit ships playable; before that plan, nothing.
          ahead: (it) => { const sent = sd.media.plans.get(it.unit); return !!sent && sent.has(it.mediaId) && sent.get(it.mediaId) === undefined },
          part: async (partPath) => {
            if (!ledger.touch(docRef, policy, sd.doc.bytes.length)) throw new Error('document cache full')
            const r = await o.pool.mediaPart(s.sid, docRef, policy, partPath, svc.temp.dir)
            return svc.temp.adopt(r.path, r.bytes)
          },
        })
      } catch { return mediaDeny(ws, reqId, 'media-unavailable') }
      sd.media.queue = q
      // A revocation of the document wakes every delivery parked in its high-water wait (fix round 1).
      if (!sd.media.revocationHooked) {
        sd.media.revocationHooked = true
        void svc.cache.revocation(docRef).then(() => { for (const w of [...sd.media.wakes]) w() })
      }
    }
    sd.media.cursor = unit
    const started = q.start(unit)
    sd.media.started = started.then(() => {}, () => {})
    let queued: number
    try { queued = await started } catch { return textAlive(ws, s.sid, sd)() ? mediaDeny(ws, reqId, 'policy-refused') : undefined }
    if (!textAlive(ws, s.sid, sd)() || sd.media.queue !== q) return
    send(ws, { reqId, type: 'media-show-ok', queued })
  }

  /** `media-cursor {unit}` (D1): no reply; ignored without a queue or with a malformed unit. */
  function onMediaCursor(ws: WS, msg: Record<string, unknown>) {
    const sid = ws.data.sid, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !sid || !sd || !owns(ws, sid) || !sd.media.queue) return
    let unit: number
    try { unit = parseMediaCursor(msg).unit } catch { return }
    if (unit < sd.doc.units.length) { sd.media.cursor = unit; sd.media.queue.cursor(unit) }
  }

  /** `media-end {}` (D1): no reply. Ends the session's queue; its cache entries stay. */
  function onMediaEnd(ws: WS, msg: Record<string, unknown>) {
    const sid = ws.data.sid, sd = sid ? sessionDoc.get(sid) : undefined
    if (ws.data.phase !== 'stream' || !sid || !sd || !owns(ws, sid)) return
    try { parseMediaEnd(msg) } catch { return }
    endMedia(sd)
  }

  /**
   * `media-request {reqId, token, mediaId}` (Task 9 order):
   * 1. the token; 2. `can.media`; 3. the show is started (`media-show`); 4. the mediaId is one of this session's
   * uses (else `invalid-media`: another session's id, a cache key, anything else; nothing is charged, recorded or
   * looked up); 5. the `'media'` rate (refuse-only); 6. the record's own plan refusal, else `queue.request`;
   * 7. the durable `media_deliveries` row (a delivery or a refusal); 8. `media-start`; 9. the chunks, each under
   * `PRINT_SEND_HIGH_WATER`, while the socket owns the session, the queue is the same and the cache still holds the
   * file (a revoked document: `media-refused media-unavailable`). An ok answer's pin is released exactly once.
   */
  async function onMediaRequest(ws: WS, msg: Record<string, unknown>) {
    const a = await mediaOp(ws, msg)
    if (!a) return
    const { reqId, s, sd } = a
    let mediaId: string, asked: number | undefined
    try { ({ mediaId, unit: asked } = parseMediaRequest(msg)) } catch { return mediaDeny(ws, reqId, 'policy-refused') }
    const deny = (reason: MediaReason) => mediaDeny(ws, reqId, reason, mediaId)
    const svc = services
    if (!svc) return deny('media-unavailable')
    const q = sd.media.queue
    if (!q) return deny('invalid-media')
    await sd.media.started
    const revoked = () => svc.cache.revoked(sd.doc.docRef)
    // Socket-and-queue liveness only (fix round 2): a dead socket or an ended queue sends nothing, while a revoked
    // document is checked on its own and answers media-refused media-unavailable.
    const alive = () => textAlive(ws, s.sid, sd)() && sd.media.queue === q
    if (!alive()) return
    // The use on the slide named by `unit`, else on the cursor's slide, else the next slide in show order using it
    // (a look-ahead request). A refusal is per (mediaId, slide): fix round 1.
    const byUnit = sd.media.uses.get(mediaId)
    const n = sd.doc.units.length, from = sd.media.cursor ?? 0
    let use: MediaUse | undefined
    if (asked !== undefined) use = byUnit?.get(asked)
    else for (let i = 0; i < n && byUnit && !use; i++) use = byUnit.get((from + i) % n)
    if (!use) return deny('invalid-media')
    // 5. Refuse-only: never a breach toward 4008, apart from the unit budget.
    const rated = ws.data.rates.check('media', now()) === 'ok'
    const key = use.item?.key
    const row = (r: { reason?: string; outBytes?: number; durationMs?: number; cacheState?: 'hit' | 'joined' | 'miss'; ffmpegVersion?: string; encodeParamsSha?: string }): boolean => {
      try {
        audit.recordMedia({ sid: s.sid, unitIndex: use.unit, mediaId, kind: 'audio', ...(key ? { sourceHash: key.sourceHash, ffmpegVersion: key.ffmpegVersion, encodeParamsSha: key.paramsSha } : {}), ...r })
        return true
      } catch { return false }
    }
    const refused = (reason: MediaReason) => (row({ reason }) ? deny(reason) : deny('policy-refused'))
    if (!rated) return refused('rate-limited')
    // Task 11b fix round: only what a plan sent for the use's unit ships playable; its refusal is the plan's own.
    const sent = sd.media.plans.get(use.unit)
    if (!sent || !sent.has(mediaId)) return refused('media-unavailable')
    const byPlan = sent.get(mediaId)
    if (byPlan) return refused(byPlan)
    if (use.refused || !use.item || !key) return refused(use.refused ?? 'media-unavailable')
    const r = await q.request(mediaId)
    if (!r.ok) return alive() ? refused(revoked() ? 'media-unavailable' : r.reason) : undefined
    // The answer is pinned for this session: released exactly once below, unless the queue has ended (dropSession
    // already released it, and a late release could hit a later show's references).
    let fh: Awaited<ReturnType<typeof openFile>> | undefined
    let wake: (() => void) | undefined
    mediaDeliveries++
    try {
      if (!alive()) return
      if (revoked()) return refused('media-unavailable')
      try { fh = await openFile(r.path, 'r') } catch { return alive() ? refused('media-unavailable') : undefined }
      const bytes = r.result.bytes
      if ((await fh.stat()).size !== bytes || bytes < 1 || bytes > MEDIA_WIRE_LIMITS.audioBytes) return refused('media-unavailable')
      await o.mediaTestHooks?.afterOpen?.()
      if (!alive()) return
      if (revoked()) return refused('media-unavailable')
      if (!row({ outBytes: bytes, durationMs: r.result.durationMs, cacheState: r.origin, ffmpegVersion: r.result.ffmpegVersion, encodeParamsSha: r.result.paramsSha })) return deny('policy-refused')
      const chunks = Math.ceil(bytes / MEDIA_WIRE_LIMITS.sendChunk)
      send(ws, { reqId, type: 'media-start', mediaId, contentType: 'audio/mp4', durationMs: r.result.durationMs, bytes, chunks })
      const job: PrintJob = { reqId, stopped: false }
      wake = () => { job.wake?.() }
      sd.media.wakes.add(wake)
      for (let i = 0; i < chunks; i++) {
        await belowHighWater(ws, job, () => alive() && !revoked())
        if (!alive()) return
        if (revoked()) return deny('media-unavailable')
        const now = svc.cache.peek(key)
        if (!now || !('path' in now) || now.path !== r.path) return deny('media-unavailable')
        const n = Math.min(MEDIA_WIRE_LIMITS.sendChunk, bytes - i * MEDIA_WIRE_LIMITS.sendChunk)
        const buf = new Uint8Array(n)
        const got = await fh.read(buf, 0, n, i * MEDIA_WIRE_LIMITS.sendChunk)
        await o.mediaTestHooks?.afterRead?.(i)
        if (!alive()) return
        if (revoked() || got.bytesRead !== n) return deny('media-unavailable')
        const f = frameChunk(reqId, buf)
        if (!f || !sendFrames(ws, [f])) return
      }
    } finally {
      mediaDeliveries--
      if (wake) sd.media.wakes.delete(wake)
      await fh?.close().catch(() => {})
      if (sd.media.queue === q && sessionDoc.get(s.sid) === sd) svc.cache.delivered(s.sid, key)
    }
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
   * the scratch lease, on the basis keyed by `label` (a unit's `${i}@${scale}`,
   * or `notesBasisLabel`). Unmarkable units ship canonical bytes with a reason. */
  async function markUnit<U extends { png: Uint8Array; w: number; h: number }>(sid: string, sd: { doc: DocState; forensic?: Issued }, label: string, unit: U, signal?: AbortSignal): Promise<{ unit: U; unmarked?: string } | RefusalCode> {
    let release: () => void
    try { release = await leaseFor(sid).acquire() } catch (e) { return e instanceof LeaseFull ? 'rate-limited' : 'policy-refused' }
    try {
      const kBasis = deriveBasisKey(o.keys.kBasisRoot, sd.doc.docRef, label)
      const r = await o.pool.encode(sid, { png: unit.png, w: unit.w, h: unit.h, kBasis, code: sd.forensic!.code, ...(o.encodeHangMs ? { hangMs: o.encodeHangMs } : {}) }, encodeEstimate(unit.w, unit.h, unit.png.length), signal)
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
    if (parseFeatures(m.features) === 'invalid') return false
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
    // Stage 4a (R6): an XLSX tiles session serves no unit or Stage 3 text op (each would render units); print works.
    const mode = ws.data.sid ? sessionDoc.get(ws.data.sid)?.sheetMode : undefined
    if (mode === 'tiles' && TILES_REFUSED.has(m.type as string)) return refuse(ws, m.reqId, 'unsupported-op')
    // Stage 4b (P1): the sheet-view ops need `sheet-view@1` (units mode and a 4a tiles session refuse them).
    if (SHEET_VIEW_OPS.has(m.type as string) && !sessionDoc.get(ws.data.sid ?? '')?.sheetView) {
      // Controller ruling: an off-mode find, copy or a11y is charged one 'text' check first (refuse-only, as Stage 3's text ops); fit and resize are not charged.
      if (SHEET_TEXT_OPS.has(m.type as string) && ws.data.rates.check('text', now()) !== 'ok') return refuse(ws, m.reqId, 'rate-limited')
      return refuse(ws, m.reqId, 'unsupported-op')
    }
    // The tile ops (never served in units mode): `sheet-layout` (Task 6), `tiles` and `sheet-cancel` (Task 9).
    if (m.type === 'tiles' || m.type === 'sheet-layout' || m.type === 'sheet-cancel') {
      if (mode !== 'tiles') return refuse(ws, m.reqId, 'unsupported-op')
      if (m.type === 'sheet-layout') return onSheetLayout(ws, m)
      if (m.type === 'tiles') return onTiles(ws, m)
      return onSheetCancel(ws, m)
    }
    // Stage 4b: a sheet-view session's ops (the gate above admits only those).
    if (m.type === 'sheet-resize') return onSheetResize(ws, m)
    if (m.type === 'sheet-fit') return onSheetFit(ws, m)
    if (m.type === 'renew-ack') return onRenewAck(ws, m)
    if (m.type === 'unit') return onUnit(ws, m)
    if (m.type === 'search') return onSearch(ws, m)
    if (m.type === 'select') return onSelect(ws, m)
    if (m.type === 'copy') return onCopy(ws, m)
    if (m.type === 'a11y') return onA11y(ws, m)
    if (m.type === 'print') return onPrint(ws, m)
    if (m.type === 'print-cancel') return onPrintCancel(ws, m)
    if (m.type === 'notes') return onNotes(ws, m)
    if (m.type === 'show-plan') return onShow(ws, m, 'plan')
    if (m.type === 'show-track') return onShow(ws, m, 'track')
    if (m.type === 'show-morph') return onShow(ws, m, 'morph')
    if (m.type === 'media-show') return onMediaShow(ws, m)
    if (m.type === 'media-request') return onMediaRequest(ws, m)
    if (m.type === 'media-cursor') return onMediaCursor(ws, m)
    if (m.type === 'media-end') return onMediaEnd(ws, m)
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
      const data: Conn = { ip, origin, url: req.url, headers: req.headers, phase: 'hello', rates: new SessionRates(o.rates), timers: new Set(), missingFaces: [], drainWaiters: new Set(), gone: new AbortController() }
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
        ws.data.gone.abort()
        wakeDrain(ws)
        for (const t of ws.data.timers) clearTimeout(t)
        const sid = ws.data.sid
        // A server-initiated close set the phase first; only a peer drop reaches here open.
        const serverClosed = ws.data.phase === 'closed'
        ws.data.phase = 'closed'
        if (!sid) return
        const s = table.get(sid)
        if (!s || s.activeSocket !== (ws as unknown)) return
        // F21: the one `'sheet'` op in flight dies with its (active) socket; a superseded socket's close leaves the new one's flag alone.
        ws.data.rates.sheetBusy = false
        // D1: a socket close is an implicit media-end, a detach included (a resume sends media-show again).
        endMedia(sessionDoc.get(sid))
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
    /** The sid's hidden slide units (tests). */
    hiddenOf: (sid: string) => sessionDoc.get(sid)?.doc.hidden,
    /** TEST-ONLY: the sid's session clock (Stage 4a, R9). */
    sheetClockOf: (sid: string) => sessionDoc.get(sid)?.clock,
    /** TEST-ONLY: whether the sid serves the 4b sheet-view ops (Stage 4b, P1). */
    sheetViewOf: (sid: string) => sessionDoc.get(sid)?.sheetView,
    /** TEST-ONLY: the sid's per-session rate state (Stage 4b, F21). */
    sheetRatesOf: (sid: string) => sessionRates.get(sid),
    /** The sid's Layer-1 issuance (tests and the validation script). */
    forensicOf: (sid: string) => sessionDoc.get(sid)?.forensic,
    /** TEST-ONLY view: the canonical cached PNG this session's raster unit is built from. */
    canonicalOf: (sid: string, unitIndex: number, scale: number) => {
      const sd = sessionDoc.get(sid)
      if (!sd) return undefined
      const hit = renderCache.get(renderKey(keyFields(sd.doc, markOf(sd.policy, sd.decision), unitIndex, scale, 'final-raster'))) as RenderedUnit | undefined
      return hit?.kind === 'raster' ? hit.png : undefined
    },
    table,
    stats: () => ({ sockets: sockets.size, sessions: table.size, docs: docs.size, cache: { entries: renderCache.count, bytes: renderCache.size }, pool: o.pool.stats(), forensic: { ...forensicStats, carriers: [...forensicStats.carriers] }, media: { deliveries: mediaDeliveries } }),
  }
  return server
}
