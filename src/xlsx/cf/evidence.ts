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

export function withCfEvidence<T>(
  capabilityId: string,
  evaluateOrVariant: string | (() => T),
  evaluate?: () => T,
): CfOutcome<T> {
  const [variant, fn] = typeof evaluateOrVariant === 'function'
    ? [undefined, evaluateOrVariant]
    : [evaluateOrVariant, evaluate!]
  return hasEvidence(capabilityId, variant)
    ? { kind: 'ready', value: fn() }
    : { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
}
