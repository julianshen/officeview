import type { CfOutcome } from './model'
import { VERIFIED_CAPABILITIES } from './verifiedCapabilities'

export { VERIFIED_CAPABILITIES }

let testInjectedEvidence: ReadonlySet<string> | null = null

/**
 * Injects evidence for testing without mutating VERIFIED_CAPABILITIES.
 * Pass null or undefined to clear injected evidence.
 */
export function injectCfEvidence(capabilities: Iterable<string> | null | undefined): void {
  testInjectedEvidence = capabilities ? new Set(capabilities) : null
}

/**
 * Scoped evidence injection helper for testing.
 */
export function withInjectedCfEvidence<T>(capabilities: Iterable<string>, fn: () => T): T {
  const previous = testInjectedEvidence
  testInjectedEvidence = new Set(capabilities)
  try {
    const result = fn()
    if (result && typeof (result as unknown as Promise<unknown>).then === 'function') {
      return (result as unknown as Promise<unknown>).finally(() => {
        testInjectedEvidence = previous
      }) as T
    }
    testInjectedEvidence = previous
    return result
  } catch (error) {
    testInjectedEvidence = previous
    throw error
  }
}

/**
 * Runtime evidence check keyed by catalog capabilityId (Ruling C1).
 * Derive verified set from manifest/oracle via auditStage1Manifest.
 */
export function hasEvidence(idOrCapability: string, variant?: string): boolean {
  if (testInjectedEvidence) {
    if (variant !== undefined) {
      if (
        testInjectedEvidence.has(`${idOrCapability}:${variant}`) ||
        testInjectedEvidence.has(variant) ||
        testInjectedEvidence.has(idOrCapability)
      ) {
        return true
      }
    } else if (testInjectedEvidence.has(idOrCapability)) {
      return true
    }
  }
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
