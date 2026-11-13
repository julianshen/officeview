/** In-repo WS harness: a real Bun.serve on port 0 + a frame-decoding client. */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { Reassembler } from '../../src/stream/frame'
import { decodeUnit } from '../../src/stream/unit'
import { validateManifest, type Manifest } from '../../src/stream/manifest'
import { deriveKeys } from '../src/keys'
import type { AuthDecision, Authorize, ResolveDocument } from '../src/hooks'
import { Pool } from '../src/pool'
import { createStreamServer, type StreamServerOptions } from '../src/ws'

export const ORIGIN = 'http://localhost:5173'
export const corpus = (name: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, '../../corpus', name)))

let sharedPool: Pool | undefined
export const pool = () => (sharedPool ??= new Pool({ workers: 2 }))
/** Per-file teardown: the next file gets a fresh pool. */
export const disposePool = () => { sharedPool?.dispose(); sharedPool = undefined }

export interface Harness { url: string; stream: ReturnType<typeof createStreamServer>; stop(): void }
export async function start(o: { policy?: AuthDecision['policy']; viewerId?: string; doc?: Uint8Array | null; authorize?: Authorize; resolveDocument?: ResolveDocument } & Partial<StreamServerOptions> = {}): Promise<Harness> {
  const keys = await deriveKeys(new Uint8Array(32).fill(9))
  const doc = o.doc === undefined ? corpus('pypptx-test.pptx') : o.doc
  const stream = createStreamServer({
    keys, originAllowlist: [ORIGIN], pool: pool(),
    authorize: o.authorize ?? (async () => ({ viewerId: o.viewerId ?? 'alice', policy: o.policy ?? {} })),
    resolveDocument: o.resolveDocument ?? (async () => (doc ? { bytes: doc, name: 'doc' } : null)),
    ...o,
  })
  const srv = Bun.serve({ port: 0, fetch: stream.fetch as never, websocket: stream.websocket as never })
  return { url: `ws://localhost:${srv.port}/stream`, stream, stop: () => srv.stop(true) }
}

export type Msg = { text: Record<string, any>; /** UTF-8 bytes as received. */ bytes?: number } | { binary: { kind: number; reqId: number; payload: Uint8Array } }
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 8: [8 << 20, 64 << 20] }

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
  async hello(extra: object = {}) {
    await this.opened()
    this.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], ...extra })
    const m = await this.next((m) => 'binary' in m && m.binary.kind === 1)
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
  close() { this.ws.close() }
}
