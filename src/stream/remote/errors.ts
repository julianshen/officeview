/** Typed stream-session failures (§6.2 rule 8 registry, client view; §10). */
import type { OperationRefusal } from '../../xlsx/view/contracts'

/** The server's session CLOSE numbers (4007 is client-originated). */
export const CLOSE_NAMES = {
  4000: 'unsupported-version', 4001: 'unsupported-format', 4002: 'ready-timeout', 4003: 'auth-failed',
  4004: 'session-expired', 4005: 'revoked', 4006: 'layout-changed', 4007: 'oversized-frame',
  4008: 'rate-limited', 4009: 'policy-refused', 4409: 'superseded',
} as const
export type CloseCodeName = (typeof CLOSE_NAMES)[keyof typeof CLOSE_NAMES]
export type StreamErrorCode =
  | CloseCodeName
  /** An op refusal that the session survives (the unit request fails). */
  | 'session-rotated'
  | 'manifest-invalid' | 'connect-failed' | 'unit-failed' | 'closed'
  /** Text-op refusals (Stage 3a): the service is off, or a copy cap was hit. */
  | 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted'
  /** Print refusals (Stage 3b): a job over the per-job cap (or too large to
   * plan, or over the client's blob ceiling: no `limit`), or over the session total. */
  | 'print-too-large' | 'print-quota-exhausted'
  /** Notes refusal (Stage 3c): the unit is out of range or hidden. */
  | 'invalid-unit'
  /** Client refusal (Stage 3c show): the slideshow workspace has no room for the surface. */
  | 'workspace-full'
  /** Show-op refusals (piece 4, D1): the slide plays its piece-3 still (`morph-fallback`: the fade). */
  | 'still' | 'show-budget' | 'show-limits' | 'show-plan-timeout' | 'morph-fallback'
  /** Client refusal (piece 4): admission had no room for a show message in flight; transient and local, never a server rate pause. */
  | 'admission'
  /** Stage 4a tile-op refusals: the op is not served in this sheet mode, the view generation is not the sheet's, or
   * the sheet's layout is over an R14 cap. */
  | 'unsupported-op' | 'stale-view' | 'layout-too-large'
  /** Stage 4b sheet-op refusals (P12): the row or column is collapsed, the sheet's coverage is unknown under
   * `visible-only`, or the session's view state is at its cap. */
  | 'hidden-target' | 'coverage-unknown' | 'view-state-limit'
  /** Stage 4c sort-op refusals (S4 §6.6; P4, P8, P13): a value-list id or cursor the server no longer holds, a typed
   * condition this viewing mode refuses, the local engine's refusal (`StreamSessionError.refusal`), a cancelled op. */
  | 'stale-filter-values' | 'filter-condition-refused' | 'view-refused' | 'cancelled'

export class StreamSessionError extends Error {
  override readonly name = 'StreamSessionError'
  /** The cap a refusal states (copy caps), when the server gave one. */
  readonly limit?: number
  /** What the refused request asked for against `limit`, when the server gave it. */
  readonly requested?: number
  /** Stage 4c (P13): a `view-refused` refusal's local engine refusal, rebuilt with local message strings. */
  readonly refusal?: OperationRefusal
  constructor(readonly code: StreamErrorCode, message: string = code, nums: { limit?: number; requested?: number; refusal?: OperationRefusal } = {}) {
    super(message)
    if (nums.limit !== undefined) this.limit = nums.limit
    if (nums.requested !== undefined) this.requested = nums.requested
    if (nums.refusal !== undefined) this.refusal = nums.refusal
  }
}

export const closeName = (code: number): CloseCodeName | undefined => (CLOSE_NAMES as Record<number, CloseCodeName>)[code]
