import { describe, expect, it } from 'bun:test'
import { AuditTable, AUDIT_RETENTION_SEC } from '../src/audit'
import { effectivePolicy } from '../src/hooks'
import { SessionTable } from '../src/sessionTable'

const sock = () => { const calls: Array<[number?, string?]> = []; return { calls, close: (c?: number, r?: string) => { calls.push([c, r]) } } }
const base = (s = sock()) => ({ sid: 's1', docRef: 'd', role: 'audience' as const, effectivePolicy: effectivePolicy({ policy: {} }), activeSocket: s, jti: 'j1', startedAt: 0 })

describe('session table', () => {
  it('rotation: both jtis valid while RENEW is outstanding, only the new after ack', () => {
    const t = new SessionTable(); t.create(base())
    t.beginRotation('s1', 'j2')
    expect([t.acceptJti('s1', 'j1'), t.acceptJti('s1', 'j2')]).toEqual(['current', 'rotating'])
    expect(t.ack('s1', 'j1')).toBe(false)            // old jti cannot ack
    expect(t.ack('s1', 'j2')).toBe(true)
    expect([t.acceptJti('s1', 'j1'), t.acceptJti('s1', 'j2'), t.acceptJti('s1', 'zz')]).toEqual(['rotated', 'current', 'unknown'])
    expect(t.ack('s1', 'j2')).toBe(false)            // no rotation pending
  })
  it('resume supersedes the prior socket with 4409 exactly once and rotates the jti', () => {
    const t = new SessionTable(), a = sock(), b = sock()
    t.create(base(a))
    t.supersede('s1', b, 'j9')
    expect(a.calls).toEqual([[4409, 'superseded']])
    expect(b.calls).toEqual([])
    expect(t.get('s1')!.activeSocket).toBe(b)
    expect(t.acceptJti('s1', 'j1')).toBe('rotated')
    expect(t.acceptJti('s1', 'j9')).toBe('current')
  })
  it('close evicts the run row', () => {
    const t = new SessionTable(); t.create(base()); t.close('s1')
    expect(t.get('s1')).toBeUndefined(); expect(t.acceptJti('s1', 'j1')).toBe('unknown')
  })
})

describe('effective policy (§2 item 5 defaults; Stage-2a forensic)', () => {
  it('defaults are protected and apply the forensic mark', () => {
    expect(effectivePolicy({ policy: {} })).toEqual({ contentPolicy: 'visible-only', projection: 'final-raster', visibleMark: 'fixed', forensicMark: true, forensicRequested: true, print: true, role: 'audience', relaxed: false })
  })
  it('relaxations are flagged', () => {
    expect(effectivePolicy({ policy: { projection: 'relaxed' } }).projection).toBe('commands')
    expect(effectivePolicy({ policy: { projection: 'relaxed' } }).relaxed).toBe(true)
    expect(effectivePolicy({ policy: { contentPolicy: 'full' } }).relaxed).toBe(true)
    expect(effectivePolicy({ policy: { visibleMark: 'viewer' } }).relaxed).toBe(false)
    expect(effectivePolicy({ policy: { print: false } }).print).toBe(false)
    // Unknown values fall back to the protected default (fail-closed).
    expect(effectivePolicy({ policy: { projection: 'wide-open' as never, contentPolicy: 'all' as never } })).toMatchObject({ projection: 'final-raster', contentPolicy: 'visible-only' })
  })
})

describe('audit table', () => {
  const row = { ip: '1.2.3.4', startedAt: 0, keyGeneration: 1, policyId: 'visible-only@1', markSettings: { visible: 'off' as const, forensic: false, forensicRequested: false }, relaxed: true, forensic: null, docRef: 'd', contentPolicy: 'visible-only' as const, visibleMark: { kind: 'off' as const } }
  it('is write-once', () => {
    const a = new AuditTable(); a.record('s1', row)
    expect(() => a.record('s1', row)).toThrow(/write-once/)
    expect(Object.isFrozen(a.get('s1'))).toBe(true)
  })
  it('expires rows after 90 days', () => {
    const a = new AuditTable(); a.record('s1', row); a.record('s2', { ...row, startedAt: 100 })
    expect(a.expire(AUDIT_RETENTION_SEC)).toBe(1)
    expect(a.get('s1')).toBeUndefined(); expect(a.get('s2')).toBeDefined()
  })
})
