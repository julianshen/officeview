import type { CfOutcome } from './model'
import { VERIFIED_CAPABILITIES } from './verifiedCapabilities'

export { VERIFIED_CAPABILITIES }

/**
 * Runtime evidence check keyed by catalog capabilityId (Ruling C1).
 * Derive verified set from manifest/oracle via auditStage1Manifest.
 */
export function hasEvidence(idOrCapability: string, variant?: string): boolean {
  if (variant !== undefined) {
    return VERIFIED_CAPABILITIES.has(`${idOrCapability}:${variant}`)
      || VERIFIED_CAPABILITIES.has(variant)
      || VERIFIED_CAPABILITIES.has(idOrCapability)
  }
  return VERIFIED_CAPABILITIES.has(idOrCapability)
}

export function withCfEvidence<T>(capabilityId: string, evaluate: () => T): CfOutcome<T> {
  return hasEvidence(capabilityId)
    ? { kind: 'ready', value: evaluate() }
    : { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
}
