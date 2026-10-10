/**
 * wildcard-pure-prototype-r3 — pure bounded ASCII wildcard matcher.
 *
 * Official documented semantics (support.microsoft.com
 * "Using wildcard characters in searches", fetched HTTP 200 this session):
 *  - ?  = any single character
 *  - *  = any number of characters (including none)
 *  - ~ followed by ?, * or ~ = that literal character
 * Undocumented tilde forms (dangling ~, ~ + other) are refused as an explicit
 * capability gate — no invented rule.
 *
 * Implementation: hand-written rolling-row dynamic-programming glob matcher.
 *  - NO RegExp is constructed from the pattern (no regex injection, no
 *    catastrophic backtracking; regex metacharacters are plain literals).
 *  - Worst-case work is (patternTokens+1) * (textLength+1) DP transitions —
 *    polynomial, no exponential blowup. No universal O(n) claim is made.
 *  - ASCII-only: any non-ASCII code unit (>= 0x80) yields an explicit
 *    unsupported-non-ascii result (never faked Unicode case folding).
 *  - caseInsensitive (default true) folds only A-Z/a-z (documented Excel
 *    text-filter case-insensitivity; ASCII scope of this helper).
 */
import {
  type WildcardOptions,
  type WildcardResult,
} from './wildcard-contracts'

const DEFAULT_MAX_PATTERN_LENGTH = 8192
const DEFAULT_MAX_DP_CELLS = 4_000_000

type Tok =
  | { t: 'lit'; ch: string } // literal (case-folded at compare time)
  | { t: 'anyOne' }
  | { t: 'anyRun' }

function fold(ch: string): string {
  const c = ch.charCodeAt(0)
  return c >= 65 && c <= 90 ? String.fromCharCode(c + 32) : ch
}

/** Tokenize; refuses undocumented tilde forms with their position. */
export function tokenizeWildcardPattern(
  pattern: string,
): Tok[] | { tildeError: true; position: number } {
  const toks: Tok[] = []
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]
    if (ch === '~') {
      const next = pattern[i + 1]
      if (next === '?') { toks.push({ t: 'lit', ch: '?' }); i++; continue }
      if (next === '*') { toks.push({ t: 'lit', ch: '*' }); i++; continue }
      if (next === '~') { toks.push({ t: 'lit', ch: '~' }); i++; continue }
      // dangling '~' or '~' + undocumented char: rule UNKNOWN, refuse
      return { tildeError: true, position: i }
    }
    if (ch === '*') { toks.push({ t: 'anyRun' }); continue }
    if (ch === '?') { toks.push({ t: 'anyOne' }); continue }
    toks.push({ t: 'lit', ch })
  }
  return toks
}

function isAscii(s: string): boolean {
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) >= 0x80) return false
  return true
}

/**
 * QUALITY Q1-residual fix: describe an invalid budget WITHOUT invoking any
 * user-controlled conversion hook (toString / valueOf / Symbol.toPrimitive /
 * toJSON). Known primitives render safely; objects and functions get a type
 * label only.
 */
function describeInvalidBudget(v: unknown): string {
  if (v === null) return 'null'
  const t = typeof v
  if (t === 'string') return JSON.stringify(v) // JSON.stringify on a string invokes nothing
  if (t === 'number' || t === 'boolean') return String(v) // String() on real numbers/booleans invokes nothing user-controlled
  if (t === 'bigint') return `${v}n`
  if (t === 'symbol') return 'symbol'
  if (t === 'function') return 'function'
  return 'object' // objects: label only — no value conversion of any kind
}

/** Full-string match (criteria semantics), bounded policy budgets. */
export function matchWildcard(
  pattern: string,
  text: string,
  caseInsensitive = true,
  options?: WildcardOptions,
): WildcardResult {
  if (typeof pattern !== 'string' || typeof text !== 'string') {
    return { status: 'unsupported-input', reason: 'pattern and text must be strings; no coercion is performed by this helper' }
  }
  // QUALITY Q1 fix: supplied budgets are validated as finite, valid,
  // nonnegative safe-integer policy BEFORE tokenization, scanning or DP
  // allocation. NaN/±Infinity previously bypassed the `x > budget` gates and
  // silently disabled the resource bounds; an invalid budget now produces an
  // explicit attributable resource diagnostic (policy, not an Excel rule).
  // Valid domain: undefined = default; nonnegative finite safe integer.
  // only `undefined` means "use the default"; null/other values are invalid
  const maxPatternLength = options?.maxPatternLength === undefined
    ? DEFAULT_MAX_PATTERN_LENGTH : options.maxPatternLength
  const maxDpCells = options?.maxDpCells === undefined
    ? DEFAULT_MAX_DP_CELLS : options.maxDpCells
  const validBudget = (v: unknown): v is number =>
    typeof v === 'number' && Number.isSafeInteger(v) && v >= 0
  if (!validBudget(maxPatternLength)) {
    return {
      status: 'resource-limit',
      reason: `invalid resource policy: maxPatternLength = ${describeInvalidBudget(maxPatternLength)} is not a nonnegative finite safe integer (policy domain: undefined=default or nonnegative finite safe integer)`,
      patternLength: pattern.length,
      textLength: text.length,
    }
  }
  if (!validBudget(maxDpCells)) {
    return {
      status: 'resource-limit',
      reason: `invalid resource policy: maxDpCells = ${describeInvalidBudget(maxDpCells)} is not a nonnegative finite safe integer (policy domain: undefined=default or nonnegative finite safe integer)`,
      patternLength: pattern.length,
      textLength: text.length,
    }
  }
  if (pattern.length > maxPatternLength) {
    return { status: 'resource-limit', reason: `pattern length ${pattern.length} exceeds the helper policy budget ${maxPatternLength}`, patternLength: pattern.length, textLength: text.length }
  }
  const toks = tokenizeWildcardPattern(pattern)
  if (!Array.isArray(toks)) {
    return { status: 'unsupported-tilde', reason: 'tilde escape is not one of the officially documented ~? ~* ~~ forms; the rule is UNKNOWN and is not invented', position: toks.position }
  }
  if (!isAscii(pattern)) {
    return { status: 'unsupported-non-ascii', reason: 'pattern contains a non-ASCII code unit; Unicode wildcard/case parity is not established and is not faked', side: 'pattern' }
  }
  if (!isAscii(text)) {
    return { status: 'unsupported-non-ascii', reason: 'text contains a non-ASCII code unit; Unicode wildcard/case parity is not established and is not faked', side: 'text' }
  }
  const n = text.length
  if ((toks.length + 1) * (n + 1) > maxDpCells) {
    return { status: 'resource-limit', reason: `matcher DP budget ${(toks.length + 1) * (n + 1)} cells exceeds the helper policy ${maxDpCells}`, patternLength: pattern.length, textLength: n }
  }

  // rolling-row DP over text positions: row[j] = prefix of tokens so far
  // matches text[0..j). Worst case (toks.length+1)*(n+1) transitions.
  let row = new Uint8Array(n + 1)
  row[0] = 1
  for (const tok of toks) {
    const next = new Uint8Array(n + 1)
    if (tok.t === 'anyRun') {
      next[0] = row[0]
      for (let j = 0; j < n; j++) next[j + 1] = (row[j + 1] || next[j]) ? 1 : 0
    } else {
      for (let j = 0; j < n; j++) {
        if (!row[j]) continue
        if (tok.t === 'anyOne') next[j + 1] = 1
        else {
          const tc = caseInsensitive ? fold(tok.ch) : tok.ch
          const tcRaw = tok.ch
          const tch = text[j]
          const matched = caseInsensitive
            ? fold(tcRaw) === fold(tch) || tc === fold(tch)
            : tcRaw === tch
          next[j + 1] = matched ? 1 : 0
        }
      }
    }
    row = next
  }
  return row[n] ? { status: 'matched' } : { status: 'not-matched' }
}
