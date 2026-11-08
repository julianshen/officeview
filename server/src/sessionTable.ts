/** Session run table (§6.2 rule 7, §9.3): evicted at session close. */
import type { EffectivePolicy } from './hooks'

export interface SocketLike { close(code?: number, reason?: string): void }
export type SessionStatus = 'manifest' | 'stream' | 'detached' | 'closed'
export interface Session {
  sid: string
  docRef: string
  viewerId?: string
  role: 'presenter' | 'audience'
  effectivePolicy: EffectivePolicy
  /** Undefined while detached (socket dropped, inside the resume grace). */
  activeSocket: SocketLike | undefined
  jti: string
  /** Set while a RENEW is outstanding: both jtis are accepted (§6.2 rule 6). */
  pendingJti?: string
  /** jtis superseded by an acknowledged rotation: ops on them get `session-rotated`. */
  rotated: Set<string>
  startedAt: number
  status: SessionStatus
}
export type JtiState = 'current' | 'rotating' | 'rotated' | 'unknown'

export class SessionTable {
  private rows = new Map<string, Session>()
  create(row: Omit<Session, 'rotated' | 'status'> & { status?: SessionStatus }): Session {
    if (this.rows.has(row.sid)) throw new Error(`sid ${row.sid} exists`)
    const s: Session = { ...row, rotated: new Set(), status: row.status ?? 'manifest' }
    this.rows.set(s.sid, s)
    return s
  }
  get(sid: string): Session | undefined { return this.rows.get(sid) }
  /** A RENEW was issued for `newJti`; until the ack both are valid. */
  beginRotation(sid: string, newJti: string): void {
    const s = this.must(sid)
    s.pendingJti = newJti
  }
  /** Classify a jti presented with an op. `current` = the session's jti;
   * `rotating` = the pending one during a RENEW. */
  acceptJti(sid: string, jti: string): JtiState {
    const s = this.rows.get(sid)
    if (!s) return 'unknown'
    if (jti === s.jti) return 'current'
    if (jti === s.pendingJti) return 'rotating'
    if (s.rotated.has(jti)) return 'rotated'
    return 'unknown'
  }
  /** renew-ack: only the pending jti completes a rotation. */
  ack(sid: string, jti: string): boolean {
    const s = this.must(sid)
    if (!s.pendingJti || jti !== s.pendingJti) return false
    s.rotated.add(s.jti)
    s.jti = jti
    s.pendingJti = undefined
    return true
  }
  /** Resume (§6.2 rule 7): atomically close the prior socket (4409) and bind the new one. */
  supersede(sid: string, socket: SocketLike, newJti: string): Session {
    const s = this.must(sid)
    const old = s.activeSocket
    s.activeSocket = socket
    s.rotated.add(s.jti)
    if (s.pendingJti) s.rotated.add(s.pendingJti)
    s.pendingJti = undefined
    s.jti = newJti
    s.status = 'manifest'
    if (old && old !== socket) old.close(4409, 'superseded')
    return s
  }
  /** The active socket dropped abnormally: keep the row for a resume. */
  detach(sid: string): void {
    const s = this.must(sid)
    s.activeSocket = undefined
    s.status = 'detached'
  }
  close(sid: string): void {
    const s = this.rows.get(sid)
    if (s) s.status = 'closed'
    this.rows.delete(sid)
  }
  get size() { return this.rows.size }
  private must(sid: string): Session {
    const s = this.rows.get(sid)
    if (!s) throw new Error(`no session ${sid}`)
    return s
  }
}
