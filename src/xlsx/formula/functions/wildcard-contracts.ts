/**
 * wildcard-pure-prototype-r3 — contract types (self-contained pure helper; no
 * ABI/production-path coupling, no criteria/lookup function scope, no cell
 * coercion — the caller owns all numeric/error/blank/origin decisions).
 */

export type WildcardResult =
  | { status: 'matched' }
  | { status: 'not-matched' }
  /** non-string inputs are rejected as an input gate, not an Excel rule */
  | { status: 'unsupported-input'; reason: string }
  /** explicit non-ASCII result — never faked Unicode case/width parity */
  | { status: 'unsupported-non-ascii'; reason: string; side: 'pattern' | 'text' }
  /**
   * documented tilde escapes are ~? ~* ~~; a tilde NOT followed by one of
   * those (dangling tilde, ~ + other char) has NO established official rule —
   * the helper refuses instead of inventing one (UNKNOWN stays UNKNOWN).
   */
  | { status: 'unsupported-tilde'; reason: string; position: number }
  /**
   * bounded application resource policy (THIS prototype's policy, not an
   * Excel claim): pattern/text budget or matcher DP-cell budget exceeded.
   */
  | {
      status: 'resource-limit'
      reason: string
      patternLength: number
      textLength: number
    }

export interface WildcardOptions {
  /** default policy budget: pattern length cap (this helper's policy) */
  maxPatternLength?: number
  /** default policy budget: matcher DP cells cap (pattern+1)*(text+1) */
  maxDpCells?: number
}
