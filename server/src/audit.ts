/** Audit table (§9.3/§13.4): write-once per sid, deleted on expiry.
 * Stage 2a also issues the Layer-1 token, salt and keyed message (§13.2).
 * Backed by the durable AuditStore (Stage 2b). Stage 3a adds the resolved
 * text services (open row + append-only changes) and per-copy records. */
import { extendedCodeword } from './mark/bch'
import { insertMessage } from './mark/message'
import { CODEC_ID } from './mark/params'
import { AUDIT_RETENTION_SEC, AuditStore, type CopyRecord, type DeliveryRecord, type StoredRow } from './audit-store'
import type { A11yMode, CopyPolicy, PrintPolicy } from './hooks'

export { AUDIT_RETENTION_SEC }

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
  /** Resolved text services at open (Stage 3a; R4: they do not set `relaxed`). A row stored before Stage 3a reads as off. */
  services: TextServices
  /** Every later value from a resume that changed them, in order, with when (the open row stays write-once). */
  servicesChanged: Array<TextServices & { at: number }>
  /** Copy events recorded for this sid and the paragraphs they carried (a11y records excluded). */
  copied: { events: number; paragraphs: number }
}
export interface TextServices { copy: CopyPolicy; textA11y: A11yMode; print: PrintPolicy }
export type AuditRowInput = Omit<AuditRow, 'untraceable' | 'forensicLayers' | 'delivered' | 'markSettings' | 'servicesChanged' | 'copied'> & { markSettings: Omit<AuditRow['markSettings'], 'forensic'> & { forensic?: boolean } }

export interface Issued { token24: number; salt8: number; msg40: Uint8Array; code: Int8Array }

/** Token space spent or no collision-free salt: refuse marked sessions (§13.2). */
export class ForensicExhausted extends Error { constructor(why: string) { super(`forensic issuance refused: ${why}`); this.name = 'ForensicExhausted' } }

const TOKEN_SPACE = 2 ** 24

export const msgHex = (bits: Uint8Array) => Array.from({ length: 5 }, (_, i) => bits.subarray(i * 8, i * 8 + 8).reduce((a, b) => a * 2 + b, 0).toString(16).padStart(2, '0')).join('')
const hexBits = (hex: string) => Array.from({ length: 40 }, (_, i) => (parseInt(hex.slice((i >> 3) * 2, (i >> 3) * 2 + 2), 16) >> (7 - i % 8)) & 1).join('')

export class AuditTable {
  private readonly store: AuditStore
  private readonly firstToken: number | undefined
  /** gen → live keyed messages as insertMessage's bit-string ids (a cache of the store; reloaded after expiry). */
  private used = new Map<number, Set<string>>()
  /** `firstToken` (TEST-ONLY) raises every generation's counter to at least that value before issuing. */
  constructor(o: { store?: AuditStore; firstToken?: number } = {}) {
    this.store = o.store ?? AuditStore.open(':memory:')
    this.firstToken = o.firstToken
  }
  private usedOf(gen: number): Set<string> {
    let u = this.used.get(gen)
    if (!u) this.used.set(gen, u = new Set(this.store.liveMessages(gen).map(hexBits)))
    return u
  }

  issue(kMsg: Uint8Array, keyGeneration: number, draw?: () => number): Issued {
    if (this.firstToken !== undefined) this.store.raiseToken(keyGeneration, this.firstToken)
    if (this.store.nextToken(keyGeneration) >= TOKEN_SPACE) throw new ForensicExhausted('24-bit token space exhausted')
    const token24 = this.store.takeToken(keyGeneration) // durable: never rewinds, even if the draw below fails
    let r: { salt8: number; msg40: Uint8Array }
    try { r = insertMessage(kMsg, keyGeneration, token24, this.usedOf(keyGeneration), draw) } catch { throw new ForensicExhausted('no collision-free salt') }
    return { token24, salt8: r.salt8, msg40: r.msg40, code: extendedCodeword(r.msg40) }
  }

  /** Write-once, durable before it returns. */
  record(sid: string, row: StoredRow): void {
    if (this.store.getRow(sid)) throw new Error(`audit row for ${sid} is write-once`)
    this.store.putRow(sid, row)
    if (row.forensic) this.usedOf(row.keyGeneration).add(hexBits(row.forensic.msg40))
  }
  /** A unit is about to go to `sid`'s socket: durable before the send. Throws for an unknown or expired sid
   * (no row behind it), so the caller refuses rather than sends unrecorded. */
  recordDelivery(d: Omit<DeliveryRecord, 'at'>): void {
    if (!this.store.getRow(d.sid)) throw new Error(`no audit row for ${d.sid}`)
    this.store.putDelivery({ ...d, at: Math.floor(Date.now() / 1000) })
  }
  /** One copy (or counted a11y answer) is about to go to `sid`: durable before the send. Throws for an unknown or expired sid. */
  recordCopy(r: Omit<CopyRecord, 'at'>): void {
    if (!this.store.getRow(r.sid)) throw new Error(`no audit row for ${r.sid}`)
    this.store.putCopy({ ...r, at: Math.floor(Date.now() / 1000) })
  }
  /** A resume changed `sid`'s resolved text services: durable before they take effect. Throws for an unknown or expired sid. */
  recordServices(sid: string, s: TextServices): void {
    if (!this.store.getRow(sid)) throw new Error(`no audit row for ${sid}`)
    this.store.putServices({ sid, copy: s.copy, textA11y: s.textA11y, print: s.print, at: Math.floor(Date.now() / 1000) })
  }
  /** Frozen view: issuance plus what has been delivered so far (distinct delivery records). */
  get(sid: string): AuditRow | undefined {
    const r = this.store.getRow(sid)
    if (!r) return undefined
    const ds = this.store.deliveries(sid)
    const marked = ds.filter((d) => d.status === 'l1').length
    const forensicLayers: Array<'l1'> = marked ? ['l1'] : []
    const copies = this.store.copies(sid).filter((c) => c.kind === 'copy')
    // A row stored before Stage 3a reads as off; one stored before 3b has no print record and reads as print off.
    const stored = (r as Partial<StoredRow>).services
    const services: TextServices = stored ? { ...stored, print: stored.print ?? false } : { copy: false, textA11y: 'off', print: false }
    return Object.freeze({
      ...r, visibleMark: Object.freeze({ ...r.visibleMark }), markSettings: Object.freeze({ ...r.markSettings, forensic: forensicLayers.length > 0 }), forensic: r.forensic && Object.freeze({ ...r.forensic }),
      forensicLayers: Object.freeze(forensicLayers) as Array<'l1'>, delivered: Object.freeze({ marked, unmarked: ds.length - marked }),
      untraceable: r.markSettings.visible === 'off' && forensicLayers.length === 0,
      services: Object.freeze({ ...services }),
      servicesChanged: Object.freeze(this.store.services(sid).map((x) => Object.freeze({ copy: x.copy, textA11y: x.textA11y, print: x.print, at: x.at }))) as AuditRow['servicesChanged'],
      copied: Object.freeze({ events: copies.length, paragraphs: copies.reduce((n, c) => n + c.paragraphs, 0) }),
    })
  }
  expire(nowSec: number): number {
    const n = this.store.expire(nowSec)
    if (n) this.used.clear()
    return n
  }
  liveMessages(gen: number): number { return this.store.liveMessages(gen).length }
  /** The token the next issue() on `gen` will try. */
  nextToken(gen: number): number { return this.store.nextToken(gen) }
}
