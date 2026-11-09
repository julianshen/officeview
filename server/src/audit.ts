/** Audit table (§9.3/§13.4): write-once per sid, deleted on expiry.
 * Stage 2a also issues the Layer-1 token, salt and keyed message (§13.2).
 * In-memory only: a restart forgets rows and resets the counter (2b persists). */
import { extendedCodeword } from './mark/bch'
import { insertMessage } from './mark/message'
import { CODEC_ID } from './mark/params'

export interface ForensicRecord { codec: typeof CODEC_ID; token24: number; salt8: number; /** 40-bit keyed message, hex (10 chars). */ msg40: string }

export interface AuditRow {
  viewerId?: string
  ip: string
  startedAt: number
  keyGeneration: number
  policyId: string
  /** `forensic` means applied, not requested. */
  markSettings: { visible: 'fixed' | 'viewer' | 'off'; forensic: boolean; forensicRequested: boolean }
  relaxed: boolean
  /** Issuance (write-once at HELLO): the code this sid's units carry when marked. */
  forensic: ForensicRecord | null
  /** Layers actually delivered on ≥ 1 unit (append-only, never from issuance alone). */
  forensicLayers: Array<'l1'>
  /** Raster units delivered marked / shipped unmarked (capacity, all-masked). */
  delivered: { marked: number; unmarked: number }
  /** Both marks off (§19.2): visible off and no forensic layer delivered. */
  untraceable: boolean
}
export type AuditRowInput = Omit<AuditRow, 'untraceable' | 'forensicLayers' | 'delivered' | 'markSettings'> & { markSettings: Omit<AuditRow['markSettings'], 'forensic'> & { forensic?: boolean } }

export interface Issued { token24: number; salt8: number; msg40: Uint8Array; code: Int8Array }

/** Token space spent or no collision-free salt: refuse marked sessions (§13.2). */
export class ForensicExhausted extends Error { constructor(why: string) { super(`forensic issuance refused: ${why}`); this.name = 'ForensicExhausted' } }

export const AUDIT_RETENTION_SEC = 90 * 24 * 3600
const TOKEN_SPACE = 2 ** 24

export const msgHex = (bits: Uint8Array) => Array.from({ length: 5 }, (_, i) => bits.subarray(i * 8, i * 8 + 8).reduce((a, b) => a * 2 + b, 0).toString(16).padStart(2, '0')).join('')
const hexBits = (hex: string) => Array.from({ length: 40 }, (_, i) => (parseInt(hex.slice((i >> 3) * 2, (i >> 3) * 2 + 2), 16) >> (7 - i % 8)) & 1).join('')

export class AuditTable {
  private rows = new Map<string, AuditRowInput>()
  /** sid → append-only delivery summary (bounded: counters, not per-unit rows). */
  private applied = new Map<string, { layers: Set<'l1'>; marked: number; unmarked: number }>()
  private nextToken_: number
  /** Live keyed messages, as insertMessage's bit-string ids. */
  private used = new Set<string>()
  constructor(o: { firstToken?: number; used?: readonly string[] } = {}) {
    this.nextToken_ = o.firstToken ?? 0
    for (const h of o.used ?? []) this.used.add(hexBits(h))
  }

  issue(kMsg: Uint8Array, keyGeneration: number, draw?: () => number): Issued {
    if (this.nextToken_ >= TOKEN_SPACE) throw new ForensicExhausted('24-bit token space exhausted')
    const token24 = this.nextToken_++ // never rewinds, even if the draw below fails
    let r: { salt8: number; msg40: Uint8Array }
    try { r = insertMessage(kMsg, keyGeneration, token24, this.used, draw) } catch { throw new ForensicExhausted('no collision-free salt') }
    return { token24, salt8: r.salt8, msg40: r.msg40, code: extendedCodeword(r.msg40) }
  }

  record(sid: string, row: AuditRowInput): void {
    if (this.rows.has(sid)) throw new Error(`audit row for ${sid} is write-once`)
    this.rows.set(sid, { ...row, markSettings: { ...row.markSettings }, forensic: row.forensic && { ...row.forensic } })
    this.applied.set(sid, { layers: new Set(), marked: 0, unmarked: 0 })
  }
  /** A unit was delivered to `sid`'s socket: marked with `layer`, or unmarked. No-op for an unknown or expired sid. */
  noteDelivered(sid: string, what: 'l1' | 'unmarked'): void {
    const a = this.applied.get(sid)
    if (!a) return
    if (what === 'unmarked') a.unmarked++
    else { a.marked++; a.layers.add(what) }
  }
  /** Frozen view: issuance plus what has been delivered so far. */
  get(sid: string): AuditRow | undefined {
    const r = this.rows.get(sid), a = this.applied.get(sid)
    if (!r || !a) return undefined
    const forensicLayers = [...a.layers]
    return Object.freeze({
      ...r, markSettings: Object.freeze({ ...r.markSettings, forensic: forensicLayers.length > 0 }), forensic: r.forensic && Object.freeze({ ...r.forensic }),
      forensicLayers: Object.freeze(forensicLayers) as Array<'l1'>, delivered: Object.freeze({ marked: a.marked, unmarked: a.unmarked }),
      untraceable: r.markSettings.visible === 'off' && forensicLayers.length === 0,
    })
  }
  expire(now: number): number {
    let n = 0
    for (const [sid, r] of this.rows) if (now - r.startedAt >= AUDIT_RETENTION_SEC) {
      this.rows.delete(sid); this.applied.delete(sid); n++
      if (r.forensic) this.used.delete(hexBits(r.forensic.msg40))
    }
    return n
  }
  get size() { return this.rows.size }
  get liveMessages() { return this.used.size }
  /** The token the next issue() will try. */
  get nextToken() { return this.nextToken_ }
}
