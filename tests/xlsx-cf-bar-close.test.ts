import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { auditStage3Manifest, unknownStage3Registrations } from '../scripts/cf-refs/audit'
import { CF_STAGE3_CAPABILITIES, CF_STAGE3_NONDISCRIMINATING, CF_STAGE3_UNSUPPORTED, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { isCapabilityVerified } from '../scripts/cf-refs/verify'

// Owner ruling S3-CLOSE (2026-11-13): three rare bar capabilities of measured cases are recorded as unsupported in
// Stage 3. The audit counts them refused-unsupported (never blocked, never verified); production is unchanged.

const REPO_ROOT = resolve(__dirname, '..')
const manifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const CLOSED = ['bar:composition-bar-stopIfTrue', 'bar:composition-unresolved-stub', 'bar:pairing-standalone']

describe('Stage 3 unsupported registry (Ruling S3-CLOSE)', () => {
  it('registers exactly the three ruled capabilities, each a known Stage 3 capability with a reason', () => {
    expect(Object.keys(CF_STAGE3_UNSUPPORTED).sort()).toEqual(CLOSED)
    const known = new Set(CF_STAGE3_CAPABILITIES.map(c => c.capabilityId))
    for (const id of CLOSED) {
      expect(known.has(id), id).toBe(true)
      expect(CF_STAGE3_UNSUPPORTED[id]!.trim().length, id).toBeGreaterThan(20)
      expect(CF_STAGE3_NONDISCRIMINATING[id], `${id} is not also non-discriminating`).toBeUndefined()
    }
    expect(unknownStage3Registrations()).toEqual([])
    // The guard has teeth: an unknown id, an empty reason and a double registration are all reported.
    expect(unknownStage3Registrations({ 'bar:nope': 'x' }, {})).toEqual(['bar:nope'])
    expect(unknownStage3Registrations({ 'bar:pairing-standalone': ' ' }, {})).toEqual(['bar:pairing-standalone'])
    expect(unknownStage3Registrations({ 'bar:pairing-standalone': 'x' }, { 'bar:pairing-standalone': 'y' })).toEqual(['bar:pairing-standalone'])
  })

  it('the audit reports them refused-unsupported, not blocked, and the verifier never lifts them', () => {
    const m = manifest()
    const report = auditStage3Manifest(m, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    for (const id of CLOSED) {
      expect(report.perCapability.find(p => p.capabilityId === id)?.disposition, id).toBe('refused-unsupported')
      expect(report.blockedCapabilities, id).not.toContain(id)
      expect(isCapabilityVerified(id, m, { rootDir: REPO_ROOT }), id).toBe(false)
    }
    // Their owning cases stay measured: the ruling closes the capability, not the capture.
    for (const id of CLOSED) {
      const owner = CF_STAGE3_CAPABILITIES.find(c => c.capabilityId === id)!.caseId
      expect(m.cases.find(c => c.id === owner)?.status, owner).toBe('measured')
    }
  }, 30_000)
})
