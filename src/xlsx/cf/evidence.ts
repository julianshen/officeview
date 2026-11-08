import type { CfOutcome } from './model'
/** Reviewed runtime allowlist. Task 1 has no measured captures yet (12 pending, one unsupported).
 * Add exact case/variant pairs only after independent evidence review. Never read fixture files here. */
const VERIFIED: readonly (readonly [id: string, variant: string])[] = Object.freeze([])
export function hasEvidence(id: string, variant?: string): boolean {
  return VERIFIED.some(pair => pair[0] === id && (variant === undefined || pair[1] === variant))
}
export function withCfEvidence<T>(id: string, variant: string, evaluate: () => T): CfOutcome<T> {
  return hasEvidence(id, variant) ? { kind: 'ready', value: evaluate() } : { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
}
