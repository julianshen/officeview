import type { CfOutcome } from './model'
import {
  VERIFIED_CAPABILITIES,
  VERIFIED_STAGE1_CAPABILITIES,
  VERIFIED_STAGE2_CAPABILITIES,
} from './verifiedCapabilities'

export {
  VERIFIED_CAPABILITIES,
  VERIFIED_STAGE1_CAPABILITIES,
  VERIFIED_STAGE2_CAPABILITIES,
}

let testInjectedEvidence: ReadonlySet<string> | null = null
// When true the injected set is the ONLY evidence (production set ignored), so tests can assert
// fail-closed behaviour with none/partial evidence regardless of what has been verified.
let testInjectedExact = false

/** Exact injection: the given set replaces the production evidence entirely. Pass null to clear. */
export function injectExactCfEvidence(capabilities: Iterable<string> | null | undefined): void {
  testInjectedEvidence = capabilities ? new Set(capabilities) : null
  testInjectedExact = capabilities != null
}

/** Scoped exact injection: inside fn, only `capabilities` count as evidence. */
export function withExactCfEvidence<T>(capabilities: Iterable<string>, fn: () => T): T {
  const previous = testInjectedEvidence
  const previousExact = testInjectedExact
  const restore = (): void => {
    testInjectedEvidence = previous
    testInjectedExact = previousExact
  }
  testInjectedEvidence = new Set(capabilities)
  testInjectedExact = true
  try {
    const result = fn()
    if (result && typeof (result as unknown as Promise<unknown>).then === 'function') {
      return (result as unknown as Promise<unknown>).finally(restore) as T
    }
    restore()
    return result
  } catch (error) {
    restore()
    throw error
  }
}

/**
 * Injects evidence for testing without mutating VERIFIED_CAPABILITIES.
 * Pass null or undefined to clear injected evidence.
 */
export function injectCfEvidence(capabilities: Iterable<string> | null | undefined): void {
  testInjectedEvidence = capabilities ? new Set(capabilities) : null
  testInjectedExact = false
}

/**
 * Scoped evidence injection helper for testing.
 */
export function withInjectedCfEvidence<T>(capabilities: Iterable<string>, fn: () => T): T {
  const previous = testInjectedEvidence
  const previousExact = testInjectedExact
  const restore = (): void => {
    testInjectedEvidence = previous
    testInjectedExact = previousExact
  }
  testInjectedEvidence = new Set(capabilities)
  testInjectedExact = false
  try {
    const result = fn()
    if (result && typeof (result as unknown as Promise<unknown>).then === 'function') {
      return (result as unknown as Promise<unknown>).finally(restore) as T
    }
    restore()
    return result
  } catch (error) {
    restore()
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
    if (testInjectedExact) return false
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
