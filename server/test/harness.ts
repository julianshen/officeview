/** In-repo WS harness: a real Bun.serve on port 0 + a frame-decoding client. */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { Reassembler } from '../../src/stream/frame'
import { decodeUnit } from '../../src/stream/unit'
import { validateManifest, type Manifest } from '../../src/stream/manifest'
import { decodeTileBody } from '../../src/stream/sheet-wire'
import { deriveKeys } from '../src/keys'
import type { AuthDecision, Authorize, ResolveDocument } from '../src/hooks'
import { Pool } from '../src/pool'
import { createStreamServer, type StreamServerOptions } from '../src/ws'
import { createMediaServices, type MediaServices, type MediaTestOnly } from '../src/media/services'

export const ORIGIN = 'http://localhost:5173'
export const corpus = (name: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, '../../corpus', name)))

let sharedPool: Pool | undefined
export const pool = () => (sharedPool ??= new Pool({ workers: 2 }))
/** Per-file teardown: the next file gets a fresh pool. Resolves once every worker of the disposed pool has closed, so
 * the next test file starts with none of them alive (Bun 1.4.5: a live Worker starves an awaiting test's messages). */
export const disposePool = (): Promise<void> => { const closed = sharedPool?.dispose() ?? Promise.resolve(); sharedPool = undefined; return closed }

/** `media`: the `MediaServices` this harness built (piece 5 Task 9), for the tests to read the ledgers. `stop()`
 * resolves once they are closed (callers that ignore it are unchanged). */
export interface Harness { url: string; stream: ReturnType<typeof createStreamServer>; media?: MediaServices; stop(): Promise<void> }
/**
 * Piece 5 (Task 9): `media` without `info` builds real media services (`createMediaServices`, its own temp root
 * unless `testOnly.root`) and, unless `pool` is given, a one-worker pool configured with them; both are closed by
 * `stop()`. `media` with `info` (Task 4 tests: no services) is passed through as before.
 */
export type HarnessMedia = { enabled?: boolean; testOnly?: MediaTestOnly }
export async function start(o: { policy?: AuthDecision['policy']; viewerId?: string; doc?: Uint8Array | null; authorize?: Authorize; resolveDocument?: ResolveDocument } & Omit<Partial<StreamServerOptions>, 'media'> & { media?: StreamServerOptions['media'] | HarnessMedia } = {}): Promise<Harness> {
  const keys = await deriveKeys(new Uint8Array(32).fill(9))
  const doc = o.doc === undefined ? corpus('pypptx-test.pptx') : o.doc
  let media: MediaServices | undefined, ownPool: Pool | undefined, root: string | undefined
  let streamMedia = o.media as StreamServerOptions['media']
  if (o.media && !('info' in o.media) && !('services' in o.media)) {
    const m = o.media as HarnessMedia
    const enabled = m.enabled ?? true
    if (!m.testOnly?.root) root = mkdtempSync(join(tmpdir(), 'ov-t9-'))
    media = await createMediaServices({ env: process.env, enabled, testOnly: { ...m.testOnly, ...(root ? { root } : {}) } })
    if (!o.pool) ownPool = new Pool({ workers: 1, ...(media.info ? { media: { info: media.info, dir: media.temp.dir } } : {}) })
    streamMedia = { enabled, ...(media.info ? { info: media.info } : {}), services: media }
  }
  // The shared pool outlives this server, which (unlike production: one server per pool) is not the only one using
  // it. After stop() the server's late teardowns (sockets it dropped detach, and are reaped after the grace) must
  // not `forget` a document a later server registered under the same docRef: scoped, forget is a no-op once stopped.
  let stopped = false
  const base = ownPool ?? o.pool ?? pool()
  const scoped = new Proxy(base, {
    get(target, key) {
      if (key === 'forget') return (docRef: string) => { if (!stopped) target.forget(docRef) }
      const v = Reflect.get(target, key, target)
      return typeof v === 'function' ? v.bind(target) : v
    },
  })
  const stream = createStreamServer({
    keys, originAllowlist: [ORIGIN],
    authorize: o.authorize ?? (async () => ({ viewerId: o.viewerId ?? 'alice', policy: o.policy ?? {} })),
    resolveDocument: o.resolveDocument ?? (async () => (doc ? { bytes: doc, name: 'doc' } : null)),
    ...o,
    pool: scoped,
    media: streamMedia,
  })
  const srv = Bun.serve({ port: 0, fetch: stream.fetch as never, websocket: stream.websocket as never })
  return {
    url: `ws://localhost:${srv.port}/stream`, stream, ...(media ? { media } : {}),
    async stop() {
      stopped = true
      srv.stop(true)
      await ownPool?.dispose()
      await media?.close()
      if (root) rmSync(root, { recursive: true, force: true })
    },
  }
}

export type Msg = { text: Record<string, any>; /** UTF-8 bytes as received. */ bytes?: number } | { binary: { kind: number; reqId: number; payload: Uint8Array } }
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 5: [8 << 20, 64 << 20], 6: [8 << 20, 32 << 20], 7: [8 << 20, 64 << 20], 8: [8 << 20, 64 << 20], 9: [4 << 20, 4 << 20] }

export class Client {
  ws: WebSocket
  msgs: Msg[] = []
  closed: Promise<{ code: number; reason: string }>
  private waiters: Array<() => void> = []
  private r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: (k) => CAPS[k]?.[1] ?? 0, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
  manifest?: Manifest
  constructor(url: string, origin = ORIGIN) {
    this.ws = new WebSocket(url, { headers: { Origin: origin } } as never)
    this.ws.binaryType = 'arraybuffer'
    this.closed = new Promise((ok) => this.ws.addEventListener('close', (e) => { ok({ code: e.code, reason: e.reason }); this.wake() }))
    // Serial receipt, like the real client's inbox: a fragment never reaches the
    // reassembler while the previous message is still inflating.
    let inbox = Promise.resolve()
    this.ws.addEventListener('message', (e) => {
      const step = inbox.then(async () => {
        if (typeof e.data === 'string') this.msgs.push({ text: JSON.parse(e.data), bytes: new TextEncoder().encode(e.data).length })
        else {
          const res = await this.r.push(new Uint8Array(e.data as ArrayBuffer))
          if (res && !res.ok) throw new Error(`frame rejected: ${res.reason}`)
          if (res?.ok) this.msgs.push({ binary: { kind: res.kind, reqId: res.reqId, payload: res.payload } })
        }
        this.wake()
      })
      // A rejected frame still surfaces as an uncaught error; later messages keep flowing.
      step.catch((err) => setTimeout(() => { throw err }))
      inbox = step.catch(() => {})
    })
  }
  private wake() { const w = this.waiters; this.waiters = []; for (const f of w) f() }
  opened() { return this.ws.readyState === WebSocket.OPEN ? Promise.resolve() : new Promise<void>((ok, fail) => { this.ws.addEventListener('open', () => ok()); this.ws.addEventListener('error', () => fail(new Error('ws error'))) }) }
  send(o: object) { this.ws.send(JSON.stringify(o)) }
  async next(pred: (m: Msg) => boolean, ms = 10_000): Promise<Msg> {
    const deadline = Date.now() + ms
    for (;;) {
      const i = this.msgs.findIndex(pred)
      if (i >= 0) return this.msgs.splice(i, 1)[0]!
      if (Date.now() > deadline || this.ws.readyState === WebSocket.CLOSED) throw new Error(`no matching message (have ${JSON.stringify(this.msgs.map((m) => 'text' in m ? m.text : { kind: m.binary.kind }))})`)
      await new Promise<void>((r) => { this.waiters.push(r); setTimeout(r, 50) })
    }
  }
  /** `ms`: the manifest wait, default `next`'s. */
  async hello(extra: object = {}, ms?: number) {
    await this.opened()
    this.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], ...extra })
    const m = await this.next((m) => 'binary' in m && m.binary.kind === 1, ms)
    const v = validateManifest(JSON.parse(new TextDecoder().decode((m as { binary: { payload: Uint8Array } }).binary.payload)))
    if (!v.ok) throw new Error(`bad manifest field ${v.field}`)
    this.manifest = v.manifest
    return v.manifest
  }
  jti(token = this.manifest!.session.token) { return JSON.parse(atob(token.split('.')[1]!.replace(/-/g, '+').replace(/_/g, '/'))).jti as string }
  async ready(extra: object = {}) { this.send({ reqId: 2, type: 'ready', tokenJti: this.jti(), ...extra }) }
  unit(reqId: number, unitIndex = 0, scale = 1, token = this.manifest!.session.token) { this.send({ reqId, type: 'unit', token, unitIndex, scale }) }
  /** All binary messages for reqId up to and including its unit message. */
  async unitMessages(reqId: number) {
    const out: Array<{ kind: number; payload: Uint8Array }> = []
    for (;;) {
      const m = await this.next((m) => ('binary' in m && m.binary.reqId === reqId) || ('text' in m && m.text.reqId === reqId && m.text.type === 'refused'))
      if ('text' in m) throw new Error(`refused ${m.text.code}`)
      out.push(m.binary)
      if (m.binary.kind === 3 || m.binary.kind === 4) return { messages: out, unit: decodeUnit(m.binary.payload) }
    }
  }
  /** Sends a text op with the current token (unless `msg.token` is given) and
   * collects its reqId's replies: every `partial` message, then the first
   * other one (`end`: e.g. `search-done` or `refused`). */
  async text(msg: { reqId: number; type: string } & Record<string, unknown>, o: { partial?: readonly string[]; ms?: number } = {}) {
    const partial = o.partial ?? ['search-result']
    this.send({ token: this.manifest!.session.token, ...msg })
    const parts: Array<{ text: Record<string, any>; bytes: number }> = []
    for (;;) {
      const m = await this.next((m) => 'text' in m && m.text.reqId === msg.reqId, o.ms) as { text: Record<string, any>; bytes: number }
      if (!partial.includes(m.text.type)) return { parts, end: m.text }
      parts.push(m)
    }
  }
  /** Sends `print` and collects its reqId: the `print-start`, every kind-8
   * body, and the closing text (`print-done` or `refused`) as `end`. */
  async print(reqId: number, from: number, to: number, o: { token?: string; ms?: number } = {}) {
    this.send({ reqId, type: 'print', token: o.token ?? this.manifest!.session.token, from, to })
    let start: Record<string, any> | undefined
    const pages: Uint8Array[] = []
    for (;;) {
      const m = await this.next((m) => ('text' in m && m.text.reqId === reqId) || ('binary' in m && m.binary.reqId === reqId), o.ms)
      if ('binary' in m) { if (m.binary.kind !== 8) throw new Error(`unexpected kind ${m.binary.kind} on print`); pages.push(m.binary.payload); continue }
      if (m.text.type === 'print-start' && !start) { start = m.text; continue }
      return { start, pages, end: m.text }
    }
  }
  /** Sends `notes` and collects its reqId: the `notes-start` and the kind-7
   * body, or the refusal (`{type: 'error'}`) as `end`; any other text on the
   * reqId is kept in `extra`. */
  async notes(unit: unknown, scale: unknown, o: { reqId?: number; token?: string; ms?: number } = {}) {
    const reqId = o.reqId ?? 700 + (this.notesSeq++)
    this.send({ reqId, type: 'notes', token: o.token ?? this.manifest!.session.token, unit, scale })
    let start: Record<string, any> | undefined
    const extra: Array<Record<string, any>> = []
    for (;;) {
      const m = await this.next((m) => ('text' in m && m.text.reqId === reqId) || ('binary' in m && m.binary.reqId === reqId), o.ms)
      if ('binary' in m) { if (m.binary.kind !== 7) throw new Error(`unexpected kind ${m.binary.kind} on notes`); return { reqId, start, body: m.binary.payload, extra } }
      if (m.text.type === 'notes-start' && !start) { start = m.text; continue }
      if (m.text.type === 'error') return { reqId, start, end: m.text, extra }
      extra.push(m.text)
    }
  }
  private notesSeq = 0
  /**
   * Piece 4: sends a show op (`show-plan`, `show-track`, `show-morph`) and collects its reqId: the `*-start`, then
   * every kind-6 payload up to `start.messages`, or the refusal (`{type: 'error'}`) as `end`. Any other text on the
   * reqId (`diag`) is kept in `extra`.
   */
  async show(msg: { type: string } & Record<string, unknown>, o: { reqId?: number; token?: string; ms?: number } = {}) {
    const reqId = o.reqId ?? 900 + (this.showSeq++)
    this.send({ reqId, token: o.token ?? this.manifest!.session.token, ...msg })
    let start: Record<string, any> | undefined
    const messages: Uint8Array[] = []
    const extra: Array<Record<string, any>> = []
    for (;;) {
      if (start && messages.length >= start.messages) return { reqId, start, messages, extra }
      const m = await this.next((m) => ('text' in m && m.text.reqId === reqId) || ('binary' in m && m.binary.reqId === reqId), o.ms)
      if ('binary' in m) { if (m.binary.kind !== 6) throw new Error(`unexpected kind ${m.binary.kind} on ${msg.type}`); messages.push(m.binary.payload); continue }
      if (m.text.type === `${msg.type}-start` && !start) { start = m.text; continue }
      if (m.text.type === 'error') return { reqId, start, messages, extra, end: m.text }
      extra.push(m.text)
    }
  }
  showPlan(unit: unknown, scale: unknown, motion: unknown, o: { reqId?: number; token?: string; ms?: number } = {}) { return this.show({ type: 'show-plan', unit, scale, motion }, o) }
  showTrack(unit: unknown, step: unknown, scale: unknown, motion: unknown, o: { reqId?: number; token?: string; ms?: number } = {}) { return this.show({ type: 'show-track', unit, step, scale, motion }, o) }
  showMorph(from: unknown, to: unknown, fromStep: unknown, scale: unknown, motion: unknown, o: { reqId?: number; token?: string; ms?: number } = {}) { return this.show({ type: 'show-morph', from, to, fromStep, scale, motion }, o) }
  private showSeq = 0
  /** Piece 5: `media-show` and its reply (`media-show-ok`, `media-refused`, or a shared `refused`). */
  async mediaShow(unit: unknown, o: { reqId?: number; token?: string; ms?: number } = {}) {
    const reqId = o.reqId ?? 1500 + (this.mediaSeq++)
    this.send({ reqId, type: 'media-show', token: o.token ?? this.manifest!.session.token, unit })
    return (await this.next((m) => 'text' in m && m.text.reqId === reqId, o.ms) as { text: Record<string, any> }).text
  }
  mediaCursor(unit: unknown) { this.send({ type: 'media-cursor', unit }) }
  mediaEnd() { this.send({ type: 'media-end' }) }
  /**
   * Piece 5: sends `media-request` and collects its reqId: the `media-start`, then every kind-9 payload up to
   * `start.chunks`, or a refusal (`media-refused` after the start, too, as a delivery the server stopped) as `end`.
   */
  async mediaRequest(mediaId: unknown, o: { reqId?: number; token?: string; ms?: number; send?: boolean; unit?: number } = {}) {
    const reqId = o.reqId ?? 1600 + (this.mediaSeq++)
    if (o.send !== false) this.send({ reqId, type: 'media-request', token: o.token ?? this.manifest!.session.token, mediaId, ...(o.unit !== undefined ? { unit: o.unit } : {}) })
    let start: Record<string, any> | undefined
    const chunks: Uint8Array[] = []
    for (;;) {
      if (start && chunks.length >= start.chunks) return { reqId, start, chunks }
      const m = await this.next((m) => ('text' in m && m.text.reqId === reqId) || ('binary' in m && m.binary.reqId === reqId), o.ms)
      if ('binary' in m) { if (m.binary.kind !== 9) throw new Error(`unexpected kind ${m.binary.kind} on media-request`); chunks.push(m.binary.payload); continue }
      if (m.text.type === 'media-start' && !start) { start = m.text; continue }
      return { reqId, start, chunks, end: m.text }
    }
  }
  private mediaSeq = 0
  /**
   * Stage 4a Task 6: sends `sheet-layout` and collects its reqId: every `sheet-layout` part (with its UTF-8 size as
   * received) up to the one with `last: true`, or any other answer (`refused`) as `end`.
   */
  async sheetLayout(reqId: number, sheetId: unknown, o: { token?: string; ms?: number } = {}) {
    this.send({ reqId, type: 'sheet-layout', token: o.token ?? this.manifest!.session.token, sheetId })
    const parts: Array<{ text: Record<string, any>; bytes: number }> = []
    for (;;) {
      const m = await this.next((m) => 'text' in m && m.text.reqId === reqId, o.ms) as { text: Record<string, any>; bytes: number }
      if (m.text.type !== 'sheet-layout') return { parts, end: m.text }
      parts.push(m)
      if (m.text.last === true) return { parts }
    }
  }
  /**
   * Stage 4a Task 9: sends `tiles` (`req` is the request without `reqId`, `type` and, unless given, `token`) and
   * collects its reqId: every kind-5 body, decoded (`decodeTileBody`), up to the closing text (`tiles-done` or a
   * whole-request `refused`) as `end`. `send: false` only collects (the caller sent it).
   */
  async tiles(reqId: number, req: Record<string, unknown>, o: { ms?: number; send?: boolean } = {}) {
    if (o.send !== false) this.send({ reqId, type: 'tiles', token: this.manifest!.session.token, ...req })
    const tiles: Array<ReturnType<typeof decodeTileBody>> = []
    for (;;) {
      const m = await this.next((m) => ('text' in m && m.text.reqId === reqId) || ('binary' in m && m.binary.reqId === reqId), o.ms)
      if ('binary' in m) { if (m.binary.kind !== 5) throw new Error(`unexpected kind ${m.binary.kind} on tiles`); tiles.push(decodeTileBody(m.binary.payload)); continue }
      return { tiles, end: m.text }
    }
  }
  close() { this.ws.close() }
}
