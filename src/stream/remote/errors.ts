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

export class StreamSessionError extends Error {
  override readonly name = 'StreamSessionError'
  constructor(readonly code: StreamErrorCode, message: string = code) { super(message) }
}

export const closeName = (code: number): CloseCodeName | undefined => (CLOSE_NAMES as Record<number, CloseCodeName>)[code]
