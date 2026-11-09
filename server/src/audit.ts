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
  forensic: ForensicRecord | null
  forensicLayers: Array<'l1'>
  /** Both marks off (§19.2): visible off and no forensic layer applied. */
  untraceable: boolean
}
export type AuditRowInput = Omit<AuditRow, 'untraceable'>

export interface Issued { token24: number; salt8: number; msg40: Uint8Array; code: Int8Array }

/** Token space spent or no collision-free salt: refuse marked sessions (§13.2). */
export class ForensicExhausted extends Error { constructor(why: string) { super(`forensic issuance refused: ${why}`); this.name = 'ForensicExhausted' } }

export const AUDIT_RETENTION_SEC = 90 * 24 * 3600
const TOKEN_SPACE = 2 ** 24

export const msgHex = (bits: Uint8Array) => Array.from({ length: 5 }, (_, i) => bits.subarray(i * 8, i * 8 + 8).reduce((a, b) => a * 2 + b, 0).toString(16).padStart(2, '0')).join('')
const hexBits = (hex: string) => Array.from({ length: 40 }, (_, i) => (parseInt(hex.slice((i >> 3) * 2, (i >> 3) * 2 + 2), 16) >> (7 - i % 8)) & 1).join('')

export class AuditTable {
  private rows = new Map<string, AuditRow>()
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
    const untraceable = row.markSettings.visible === 'off' && row.forensicLayers.length === 0
    this.rows.set(sid, Object.freeze({ ...row, markSettings: Object.freeze({ ...row.markSettings }), forensic: row.forensic && Object.freeze({ ...row.forensic }), forensicLayers: Object.freeze([...row.forensicLayers]) as Array<'l1'>, untraceable }))
  }
  get(sid: string): AuditRow | undefined { return this.rows.get(sid) }
  expire(now: number): number {
    let n = 0
    for (const [sid, r] of this.rows) if (now - r.startedAt >= AUDIT_RETENTION_SEC) {
      this.rows.delete(sid); n++
      if (r.forensic) this.used.delete(hexBits(r.forensic.msg40))
    }
    return n
  }
  get size() { return this.rows.size }
  get liveMessages() { return this.used.size }
  /** The token the next issue() will try. */
  get nextToken() { return this.nextToken_ }
}
