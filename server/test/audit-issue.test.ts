/** Stage 2a Task 3: token/salt/msg issuance from the audit table, keys (§13.2/§13.4). */
import { describe, expect, it } from 'bun:test'
import { AuditTable, ForensicExhausted } from '../src/audit'
import type { StoredRow } from '../src/audit-store'
import { deriveKeys } from '../src/keys'
import { extendedCodeword } from '../src/mark/bch'
import { messageFor } from '../src/mark/message'

const kMsg = Uint8Array.from({ length: 32 }, (_, i) => 200 - i)
const base: StoredRow = { ip: '1.2.3.4', startedAt: 0, keyGeneration: 1, policyId: 'visible-only@1', markSettings: { visible: 'off', forensic: false, forensicRequested: false }, relaxed: false, forensic: null, docRef: 'd', contentPolicy: 'visible-only', visibleMark: { kind: 'off' }, services: { copy: false, textA11y: 'off', print: false } }
/** A table whose generation-1 message for (token, salt) is already live. */
const withLive = (token: number, salt: number) => { const t = new AuditTable({ firstToken: token }); t.record('live', { ...base, forensic: { codec: 'l1-bch512x40@1', token24: token, salt8: salt, msg40: hexOf(messageFor(kMsg, 1, token, salt)) } }); return t }
const deliver = (t: AuditTable, sid: string, status: 'l1' | 'unmarked') => t.recordDelivery({ sid, unitIndex: status === 'l1' ? 0 : 1, scale: 1, status, ...(status === 'unmarked' ? { reason: 'capacity' } : {}), refSha256: 'r' })
const hexOf = (bits: Uint8Array) => Array.from({ length: 5 }, (_, i) => bits.subarray(i * 8, i * 8 + 8).reduce((a, b) => a * 2 + b, 0).toString(16).padStart(2, '0')).join('')

describe('AuditTable.issue', () => {
  it('issues monotonic tokens 0,1,2 with distinct keyed messages and their 512-sign code', () => {
    const t = new AuditTable()
    const a = t.issue(kMsg, 1), b = t.issue(kMsg, 1), c = t.issue(kMsg, 1)
    expect([a.token24, b.token24, c.token24]).toEqual([0, 1, 2])
    expect(hexOf(a.msg40)).not.toBe(hexOf(b.msg40))
    expect(a.msg40).toEqual(messageFor(kMsg, 1, 0, a.salt8))
    expect(a.code).toEqual(extendedCodeword(a.msg40))
  })
  it('re-draws a colliding salt and refuses a stuck RNG with ForensicExhausted', () => {
    const t = new AuditTable({ firstToken: 9 })
    // Force a collision: token 9's message with salt 0 is already live.
    const live = t.issue(kMsg, 1, () => 0)
    expect(live.token24).toBe(9)
    const t2 = withLive(9, 0)
    expect(() => t2.issue(kMsg, 1, () => 0)).toThrow(ForensicExhausted)
    const draws = [0, 0, 5]
    const ok = withLive(9, 0).issue(kMsg, 1, () => draws.shift()!)
    expect(ok.salt8).toBe(5)
  })
  it('issues token 2^24-1 once, then refuses', () => {
    const t = new AuditTable({ firstToken: 2 ** 24 - 1 })
    expect(t.issue(kMsg, 1).token24).toBe(2 ** 24 - 1)
    expect(() => t.issue(kMsg, 1)).toThrow(ForensicExhausted)
  })
  it('expire frees the row msg40 but the token counter never rewinds', () => {
    const t = new AuditTable()
    const f = t.issue(kMsg, 1, () => 3)
    t.record('s1', { ...base, forensic: { codec: 'l1-bch512x40@1', token24: f.token24, salt8: f.salt8, msg40: hexOf(f.msg40) } })
    deliver(t, 's1', 'l1')
    expect(t.liveMessages(1)).toBe(1)
    expect(t.expire(91 * 24 * 3600)).toBe(1)
    expect(t.liveMessages(1)).toBe(0)
    expect(() => deliver(t, 's1', 'l1')).toThrow(/no audit row/) // expired: refused, never recorded against nothing
    expect(t.get('s1')).toBeUndefined()
    expect(t.issue(kMsg, 1).token24).toBe(f.token24 + 1)
  })
})

describe('AuditTable.record', () => {
  it('derives untraceable from what was delivered, not what was issued', () => {
    const t = new AuditTable()
    t.record('a', base)
    expect(t.get('a')!.untraceable).toBe(true)
    // Issued but nothing delivered yet: no layer applied, still untraceable.
    t.record('b', { ...base, forensic: { codec: 'l1-bch512x40@1', token24: 0, salt8: 1, msg40: '0011223344' } })
    expect(t.get('b')).toMatchObject({ forensicLayers: [], markSettings: { forensic: false }, untraceable: true, delivered: { marked: 0, unmarked: 0 } })
    deliver(t, 'b', 'unmarked')
    expect(t.get('b')).toMatchObject({ forensicLayers: [], untraceable: true, delivered: { marked: 0, unmarked: 1 } })
    deliver(t, 'b', 'l1')
    expect(t.get('b')).toMatchObject({ forensicLayers: ['l1'], markSettings: { forensic: true }, untraceable: false, delivered: { marked: 1, unmarked: 1 } })
    expect(Object.isFrozen(t.get('b'))).toBe(true)
    t.record('c', { ...base, markSettings: { ...base.markSettings, visible: 'fixed' } })
    expect(t.get('c')!.untraceable).toBe(false)
    expect(() => deliver(t, 'gone', 'l1')).toThrow(/no audit row/) // unknown sid
  })
})

it('deriveKeys returns distinct 32-byte kBasisRoot and kMsg at key generation 1', async () => {
  const k = await deriveKeys(new Uint8Array(32).fill(7))
  expect(k.kBasisRoot.length).toBe(32)
  expect(k.kMsg.length).toBe(32)
  expect(Buffer.from(k.kBasisRoot).equals(Buffer.from(k.kMsg))).toBe(false)
  expect(Buffer.from(k.kMsg).equals(Buffer.from(k.kAudit))).toBe(false)
  expect(k.keyGeneration).toBe(1)
})
