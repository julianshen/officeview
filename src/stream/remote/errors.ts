/** Typed stream-session failures (§6.2 rule 8 registry, client view; §10). */

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

export class StreamSessionError extends Error {
  override readonly name = 'StreamSessionError'
  /** The cap a refusal states (copy caps), when the server gave one. */
  readonly limit?: number
  /** What the refused request asked for against `limit`, when the server gave it. */
  readonly requested?: number
  constructor(readonly code: StreamErrorCode, message: string = code, nums: { limit?: number; requested?: number } = {}) {
    super(message)
    if (nums.limit !== undefined) this.limit = nums.limit
    if (nums.requested !== undefined) this.requested = nums.requested
  }
}

export const closeName = (code: number): CloseCodeName | undefined => (CLOSE_NAMES as Record<number, CloseCodeName>)[code]
