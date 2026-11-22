/**
 * Stage 4a (S4 §6.4, §6.5, §6.11): the client side of the `sheet-tiles@1` ops on one stream session — `sheet-layout`
 * answers and `tiles` batches. The session owns the socket, the reassembler and resume; this module owns each
 * request's state and checks everything the server sends for it before the caller sees it.
 *
 * Fail closed against a hostile or buggy server (Review Focus 5): a kind-5 message is reassembled only for a live,
 * sent batch with tiles still pending (`expectsTile`); a completed one is handed on only when its body decodes
 * (`decodeTileBody`: header and IHDR, never the pixels) to a coordinate of that batch not yet delivered, on the
 * batch's sheet, `viewGen`, `zoomStep` and `dpr`, at 512 × 512. Anything else is dropped, and the session goes on.
 * A batch settles once, and every coordinate it asked for ends delivered or refused: `tiles-done` names the refused
 * ones, and any it leaves out are refused `policy-refused`. A resume re-sends exactly the pending coordinates, once,
 * under a fresh reqId and the new token (Review Focus 1).
 *
 * React-free, and nothing from `src/xlsx/**` at runtime (F5).
 */
import type { Manifest } from '../manifest'
import { decodeSheetLayout, decodeTileBody, LAYOUT_LIMITS, LAYOUT_PART_BYTES, TILE_BATCH_MAX, TILE_PALETTES, TILE_ZOOM_STEPS, type StreamSheetLayout, type TilePalette } from '../sheet-wire'
import { StreamSessionError, type StreamErrorCode } from './errors'
import { utf8Bytes } from './text-wire'

/** One checked tile: the request's sheet, view, zoom step and dpr, a requested coordinate, and its 512² PNG (undecoded). */
export interface DecodedTile { sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; col: number; row: number; png: Uint8Array }
export interface TileCoord { col: number; row: number }
export interface TilesRequest { sheetId: string; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; tiles: TileCoord[] }
export interface TilesResult { refused: Array<TileCoord & { code: string }>; retryAfterMs?: number; cancelled?: boolean }

/** A `rate-limited` refusal without a usable `retryAfterMs` (admission and lease refusals send none): retry after this. */
export const TILE_RETRY_FALLBACK_MS = 1_000
/** A server's `retryAfterMs` is clamped to the `'tile'` rate window (R4: 600 per 10 s). */
export const TILE_RETRY_MAX_MS = 10_000
/** A refusal code longer than this is not one (it is reported `policy-refused`). */
const CODE_CHARS = 32
/** The layout refusals that reject with their own code; any other is `policy-refused`. */
const LAYOUT_REFUSALS = new Set<StreamErrorCode>(['policy-refused', 'rate-limited', 'layout-too-large', 'unsupported-op', 'session-rotated'])

/** What the channel needs of its session. */
export interface SheetChannelHost {
  manifest(): Manifest
  /** The current session token (rotated by RENEW, reissued by resume). */
  token(): string
  allocId(): number
  /** The socket is open and the session is not resuming. */
  canSend(): boolean
  send(m: object): void
  /** Drops any partial reassembly of `reqId`. */
  dropReassembly(reqId: number): void
  /** A request went out (the session's reassembly sweeper runs while any is pending). */
  onSent(): void
}

interface Batch {
  readonly req: Readonly<Omit<TilesRequest, 'tiles'>>
  readonly sheetIndex: number
  reqId: number
  /** Sent on the current socket (false while it waits out a resume). */
  sent: boolean
  rotatedRetry: boolean
  /** The coordinates not yet delivered, in request order. */
  readonly pending: Map<string, TileCoord>
  readonly onTile: (t: DecodedTile) => void
  finish: (r: TilesResult) => void
  fail: (e: unknown) => void
}
interface LayoutReq {
  readonly sheetId: string
  reqId: number
  sent: boolean
  rotatedRetry: boolean
  /** The current issue's parts, in arrival order. */
  parts: unknown[]
  resolve: (l: StreamSheetLayout) => void
  fail: (e: unknown) => void
}

const abortError = () => Object.assign(new Error('The stream session was closed.'), { name: 'AbortError' })
const u32 = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0 && (n as number) <= 0xffffffff
const keyOf = (col: number, row: number) => `${col},${row}`
const codeOf = (c: unknown): string => typeof c === 'string' && c.length > 0 && c.length <= CODE_CHARS ? c : 'policy-refused'
/** The retry delay of a result with a `rate-limited` refusal: the server's when sane (clamped), else the fallback. */
const retryAfter = (v: unknown): number => typeof v === 'number' && Number.isFinite(v) && v > 0 ? Math.min(Math.ceil(v), TILE_RETRY_MAX_MS) : TILE_RETRY_FALLBACK_MS

export class SheetChannel {
  private readonly batches = new Map<number, Batch>()
  private readonly layouts = new Map<number, LayoutReq>()
  /** Kind-5 messages dropped by the checks (tests and diagnostics). */
  dropped = 0

  constructor(private readonly host: SheetChannelHost) {}

  /** Requests in flight or waiting out a resume. */
  get pending(): number { return this.batches.size + this.layouts.size }

  private tilesMode(): Manifest {
    const m = this.host.manifest()
    if (m.sheetMode !== 'tiles' || !m.sheets) throw new StreamSessionError('unsupported-op', 'this session is not in tiles mode')
    return m
  }

  /** See StreamSession.sheetLayout. */
  sheetLayout(sheetId: string, signal?: AbortSignal): Promise<StreamSheetLayout> {
    const m = this.tilesMode()
    if (signal?.aborted) return Promise.reject(signal.reason ?? abortError())
    if (typeof sheetId !== 'string' || !m.sheets!.some((s) => s.sheetId === sheetId)) return Promise.reject(new StreamSessionError('policy-refused', `no sheet ${String(sheetId)}`))
    return new Promise<StreamSheetLayout>((resolve, reject) => {
      let settled = false
      const settle = (done: () => void) => {
        if (settled) return
        settled = true
        if (this.layouts.get(req.reqId) === req) this.layouts.delete(req.reqId)
        signal?.removeEventListener('abort', onAbort)
        req.parts = []
        done()
      }
      const onAbort = () => req.fail(signal!.reason ?? abortError())
      const req: LayoutReq = {
        sheetId, reqId: 0, sent: false, rotatedRetry: false, parts: [],
        resolve: (l) => settle(() => resolve(l)),
        fail: (e) => settle(() => reject(e)),
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.issueLayout(req)
    })
  }

  /** See StreamSession.tiles. */
  tiles(r: TilesRequest, onTile: (t: DecodedTile) => void, signal?: AbortSignal): Promise<TilesResult> {
    const m = this.tilesMode()
    const bad = (why: string) => Promise.reject(new StreamSessionError('policy-refused', `tiles: ${why}`))
    if (!r || typeof r !== 'object') return bad('no request')
    if (typeof r.sheetId !== 'string' || !m.sheets!.some((s) => s.sheetId === r.sheetId)) return bad(`no sheet ${String(r.sheetId)}`)
    if (!u32(r.viewGen)) return bad('viewGen')
    if (!Number.isInteger(r.zoomStep) || r.zoomStep < 0 || r.zoomStep >= TILE_ZOOM_STEPS.length) return bad('zoomStep')
    if (r.dpr !== 1 && r.dpr !== 2) return bad('dpr')
    if (!(TILE_PALETTES as readonly unknown[]).includes(r.palette)) return bad('palette')
    if (!Array.isArray(r.tiles) || r.tiles.length === 0 || r.tiles.length > TILE_BATCH_MAX) return bad('tile count')
    const pending = new Map<string, TileCoord>()
    for (const t of r.tiles) {
      if (!t || !u32(t.col) || !u32(t.row)) return bad('coordinate')
      const k = keyOf(t.col, t.row)
      if (pending.has(k)) return bad('duplicate coordinate')
      pending.set(k, { col: t.col, row: t.row })
    }
    if (signal?.aborted) return Promise.reject(signal.reason ?? abortError())
    return new Promise<TilesResult>((resolve, reject) => {
      let settled = false
      const settle = (done: () => void) => {
        if (settled) return
        settled = true
        if (this.batches.get(b.reqId) === b) this.batches.delete(b.reqId)
        if (b.reqId) this.host.dropReassembly(b.reqId)
        signal?.removeEventListener('abort', onAbort)
        done()
      }
      // Abort: the server is told to start no more of the batch; nothing more of it is delivered.
      const onAbort = () => {
        if (settled) return
        if (b.sent && this.host.canSend()) { try { this.host.send({ reqId: this.host.allocId(), type: 'sheet-cancel', target: b.reqId }) } catch { /* the socket is failing */ } }
        b.finish({ refused: [], cancelled: true })
      }
      const b: Batch = {
        req: { sheetId: r.sheetId, viewGen: r.viewGen, zoomStep: r.zoomStep, dpr: r.dpr, palette: r.palette },
        sheetIndex: Number(r.sheetId), reqId: 0, sent: false, rotatedRetry: false, pending, onTile,
        finish: (x) => settle(() => resolve(x)),
        fail: (e) => settle(() => reject(e)),
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.issueTiles(b)
    })
  }

  /** Reassemble a kind-5 message for `reqId` only while it names a live, sent batch with tiles still pending. */
  expectsTile(reqId: number): boolean {
    const b = this.batches.get(reqId)
    return !!b && b.sent && b.pending.size > 0
  }

  /** One completed kind-5 message: handed on only when it is a pending tile of its batch (see the module doc). */
  onTileFrame(reqId: number, payload: Uint8Array): void {
    const b = this.batches.get(reqId)
    if (!b || !b.sent) { this.dropped++; return }
    let t: ReturnType<typeof decodeTileBody>
    try { t = decodeTileBody(payload) } catch { this.dropped++; return }
    const k = keyOf(t.col, t.row)
    if (t.sheetIndex !== b.sheetIndex || t.viewGen !== b.req.viewGen || t.zoomStep !== b.req.zoomStep || t.dpr !== b.req.dpr || !b.pending.has(k)) { this.dropped++; return }
    b.pending.delete(k)
    try { b.onTile({ sheetIndex: t.sheetIndex, viewGen: t.viewGen, zoomStep: t.zoomStep, dpr: t.dpr, col: t.col, row: t.row, png: t.png }) } catch { /* a callback never breaks the session */ }
  }

  /** A text message for a live request of this channel; false when it is not one (the session routes it on). */
  onText(m: Record<string, unknown>, data: string): boolean {
    if (!u32(m.reqId)) return false
    const b = this.batches.get(m.reqId)
    if (b?.sent) { this.onBatchText(b, m, data); return true }
    const l = this.layouts.get(m.reqId)
    if (l?.sent) { this.onLayoutText(l, m, data); return true }
    return false
  }

  /** The socket dropped: nothing in flight is live on it any more (a resume re-sends). */
  onDrop(): void {
    for (const b of this.batches.values()) b.sent = false
    for (const l of this.layouts.values()) { l.sent = false; l.parts = [] }
  }

  /** The session resumed: every request not sent on the new socket goes out once, under a fresh reqId. */
  onResumed(): void {
    for (const b of [...this.batches.values()]) {
      if (b.sent) continue
      // Every tile arrived before the drop (only `tiles-done` was missing).
      if (!b.pending.size) b.finish({ refused: [] })
      else this.issueTiles(b)
    }
    for (const l of [...this.layouts.values()]) if (!l.sent) this.issueLayout(l)
  }

  /** The session ended: every request rejects with `err`. */
  end(err: unknown): void {
    for (const b of [...this.batches.values()]) b.fail(err)
    for (const l of [...this.layouts.values()]) l.fail(err)
  }

  // ── internals ──────────────────────────────────────────────────────────

  private issueTiles(b: Batch) {
    if (b.reqId) { if (this.batches.get(b.reqId) === b) this.batches.delete(b.reqId); this.host.dropReassembly(b.reqId) }
    b.reqId = this.host.allocId(); b.sent = false
    this.batches.set(b.reqId, b)
    if (!this.host.canSend()) return
    try {
      this.host.send({ reqId: b.reqId, type: 'tiles', token: this.host.token(), ...b.req, tiles: [...b.pending.values()].map((t) => ({ col: t.col, row: t.row })) })
    } catch { return b.fail(new StreamSessionError('policy-refused', 'the tiles request could not be sent')) }
    b.sent = true
    this.host.onSent()
  }

  private issueLayout(l: LayoutReq) {
    if (l.reqId && this.layouts.get(l.reqId) === l) this.layouts.delete(l.reqId)
    l.reqId = this.host.allocId(); l.sent = false; l.parts = []
    this.layouts.set(l.reqId, l)
    if (!this.host.canSend()) return
    try {
      this.host.send({ reqId: l.reqId, type: 'sheet-layout', token: this.host.token(), sheetId: l.sheetId })
    } catch { return l.fail(new StreamSessionError('policy-refused', 'the sheet-layout request could not be sent')) }
    l.sent = true
  }

  /** A batch's whole-request refusal or its `tiles-done`. Any other message for it is ignored. */
  private onBatchText(b: Batch, m: Record<string, unknown>, data: string) {
    if (m.type === 'refused' || m.type === 'error') {
      // A RENEW race (stale token, refused before anything was charged): re-sent once.
      if (m.code === 'session-rotated' && !b.rotatedRetry) { b.rotatedRetry = true; return this.issueTiles(b) }
      const code = codeOf(m.code)
      const refused = [...b.pending.values()].map((t) => ({ col: t.col, row: t.row, code }))
      return b.finish({ refused, ...(code === 'rate-limited' ? { retryAfterMs: retryAfter(m.retryAfterMs) } : {}) })
    }
    if (m.type !== 'tiles-done') return
    // Over 64 KiB, `tiles-done` is not one: every pending tile is refused.
    const list = utf8Bytes(data) <= LAYOUT_PART_BYTES && Array.isArray(m.refused) ? m.refused as unknown[] : []
    const refused: TilesResult['refused'] = [], named = new Set<string>()
    for (const x of list.slice(0, TILE_BATCH_MAX)) {
      if (!x || typeof x !== 'object') continue
      const { col, row, code } = x as Record<string, unknown>
      if (!u32(col) || !u32(row)) continue
      const k = keyOf(col, row)
      // Only a coordinate still pending, once.
      if (!b.pending.has(k) || named.has(k)) continue
      named.add(k)
      refused.push({ col, row, code: codeOf(code) })
    }
    for (const [k, t] of b.pending) if (!named.has(k)) refused.push({ col: t.col, row: t.row, code: 'policy-refused' })
    b.finish({
      refused,
      ...(refused.some((x) => x.code === 'rate-limited') ? { retryAfterMs: retryAfter(m.retryAfterMs) } : {}),
      ...(m.cancelled === true ? { cancelled: true } : {}),
    })
  }

  /** A layout request's refusal or one of its parts. Any other message for it is ignored. */
  private onLayoutText(l: LayoutReq, m: Record<string, unknown>, data: string) {
    if (m.type === 'refused' || m.type === 'error') {
      if (m.code === 'session-rotated' && !l.rotatedRetry && !l.parts.length) { l.rotatedRetry = true; return this.issueLayout(l) }
      const code = LAYOUT_REFUSALS.has(m.code as StreamErrorCode) ? m.code as StreamErrorCode : 'policy-refused'
      return l.fail(new StreamSessionError(code))
    }
    if (m.type !== 'sheet-layout') return
    const hostile = (why: string) => l.fail(new StreamSessionError('policy-refused', why))
    if (utf8Bytes(data) > LAYOUT_PART_BYTES) return hostile('a sheet-layout part over 64 KiB')
    l.parts.push(m)
    if (l.parts.length > LAYOUT_LIMITS.parts) return hostile(`more than ${LAYOUT_LIMITS.parts} sheet-layout parts`)
    if (m.last !== true) return
    let layout: StreamSheetLayout
    try { layout = decodeSheetLayout(l.parts) } catch (e) { return hostile((e as Error).message) }
    if (layout.sheetId !== l.sheetId) return hostile(`a layout of sheet ${layout.sheetId}, asked for ${l.sheetId}`)
    l.resolve(layout)
  }
}
