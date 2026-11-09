/** Audit table (§9.3/§13.4): write-once per sid, deleted on expiry. */
export interface AuditRow {
  viewerId?: string
  ip: string
  startedAt: number
  keyGeneration: number
  policyId: string
  markSettings: { visible: 'fixed' | 'viewer' | 'off'; forensic: boolean; forensicRequested: boolean }
  relaxed: boolean
  /** Both marks off (§19.2). In Stage 1 the forensic mark is never applied, so
   * this is true whenever the visible mark is off. */
  untraceable: boolean
}

export const AUDIT_RETENTION_SEC = 90 * 24 * 3600

export class AuditTable {
  private rows = new Map<string, AuditRow>()
  record(sid: string, row: AuditRow): void {
    if (this.rows.has(sid)) throw new Error(`audit row for ${sid} is write-once`)
    this.rows.set(sid, Object.freeze({ ...row, markSettings: Object.freeze({ ...row.markSettings }) }))
  }
  get(sid: string): AuditRow | undefined { return this.rows.get(sid) }
  expire(now: number): number {
    let n = 0
    for (const [sid, r] of this.rows) if (now - r.startedAt >= AUDIT_RETENTION_SEC) { this.rows.delete(sid); n++ }
    return n
  }
  get size() { return this.rows.size }
}
