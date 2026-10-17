/**
 * format.ts — finite common number/date format grammar shared by TEXT and the
 * numeric cell renderer (C3 integration of the accepted pure helpers).
 *
 * ABI.md: formatCommon(value, format, semantics):
 *   { kind: 'text'; text: string } | { kind: 'unsupported'; reason: string } | EvaluationError
 *
 * Finite published grammar only (SEMANTICS.md finite-grammar list + measured
 * native rows): 0/# digit placeholders, decimals/grouping/scaling commas,
 * percent, quoted and backslash-escaped literal text, positive/negative/zero/
 * text sections, date/time tokens with month/minute disambiguation and AM/PM,
 * en-US month/weekday names for locale 'en-US'. Colors, conditions, currency/
 * locale prefixes, elapsed [h], scientific, '_', '?' and 'General' tokens are
 * DIAGNOSED ({kind:'unsupported'}) — never invented.
 */
import type { ElementValue, FormatResult, ResolvedSemantics } from './types'
import { isEvaluationError } from './evaluator'
import { decodeSerial, timeOfDayParts, weekdayIndex } from './serial'

/** Internal, synchronous preallocation observer used by guarded copy. */
export interface FormatAllocation {
  allocate(bytes: number): void; release(bytes: number): void; borrowed?(text: string): void
  /** Hands builders that outlive the formatter to `targets`; `kind` 'diagnostic'
   * marks a source diagnostic (it lives with the source, not the operation). */
  retain?(targets: readonly object[], bytes: number, commit: () => void, kind?: 'diagnostic'): void
  /** Output builders stop at this many UTF-16 units (chart labels); `truncated` is told when they do. */
  maxOutput?: number
  truncated?(): void
}
/** Old and new UTF-16 builders coexist at concat; tokens/source remain borrowed. */
function appendText(out: string, text: string, allocation?: FormatAllocation): string {
  allocation?.allocate(2 * (out.length + text.length))
  const next = out + text
  allocation?.release(out.length * 2)
  return next
}

/** The final emission, cut to `allocation.maxOutput` (reported as truncated). */
function capOut(text: string, allocation?: FormatAllocation): string {
  const cap = allocation?.maxOutput
  if (cap === undefined || text.length <= cap) return text
  allocation!.truncated?.()
  return text.slice(0, cap)
}
/** An output builder: like appendText, but never grows past `allocation.maxOutput`. */
function appendOut(out: string, text: string, allocation?: FormatAllocation): string {
  const cap = allocation?.maxOutput
  if (cap !== undefined) {
    const room = cap - out.length
    if (room <= 0) { if (text.length) allocation!.truncated?.(); return out }
    if (text.length > room) { allocation!.truncated?.(); text = text.slice(0, room) }
  }
  return appendText(out, text, allocation)
}

type Tok =
  | { t: 'lit'; text: string }
  | { t: 'ph'; ch: '0' | '#' } // digit placeholder
  | { t: 'dot' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'at' }
  | { t: 'am' } // AM/PM designator
  | { t: 'dt'; ch: 'y' | 'm' | 'd' | 'h' | 's'; len: number }

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December']
const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const WEEKDAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

/** Half-away-from-zero decimal rounding on the shortest decimal representation. */
export function roundHalfAwayDecimal(x: number, decimals: number, allocation?: FormatAllocation): number {
  if (!Number.isFinite(x)) return x
  if (x === 0) return 0
  const neg = x < 0
  allocation?.allocate(512) // finite Number repr (<= 25 units), captures and digit scratch
  const str = String(neg ? -x : x) // shortest repr of |x|
  const m = /^(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/.exec(str)
  if (!m) return x
  const ip = m[1]
  const fp = m[2] ?? ''
  const ex = m[3] ? Number(m[3]) : 0
  let digits = ip + fp
  let point = ip.length + ex
  // strip leading zeros (adjust point)
  let i = 0
  while (i < digits.length - 1 && digits[i] === '0' && point - i - 1 > -decimals - 1) i++
  if (i > 0 && i < digits.length) { digits = digits.slice(i); point -= i }
  // strip trailing zeros
  digits = digits.replace(/0+$/, '') || '0'
  const keep = point + decimals // digits kept; digit index `keep` is first dropped
  if (keep < 0) return 0 // SPEC F5 fix: keep==0 has its leading digit at the round boundary
  const up = keep < digits.length && digits[keep] >= '5'
  allocation?.allocate(2 * Math.max(0, keep))
  let kept = keep >= digits.length ? digits.padEnd(keep, '0') : digits.slice(0, keep)
  if (up) {
    // increment the kept digit string (BigInt for safety)
    // Two BigInt magnitudes plus decimal output fit four bytes per digit.
    allocation?.allocate(4 * (kept.length + 1))
    kept = (BigInt(kept) + 1n).toString()
  }
  // value = int(kept) * 10^(point - keep) = int(kept) * 10^(-decimals)
  const mag = stringTimesPow10(kept, point - keep, allocation)
  return neg ? -mag : mag
}

/** exact decimal-string times 10^e, converted to Number (lossless for modest sizes). */
function stringTimesPow10(intStr: string, e: number, allocation?: FormatAllocation): number {
  allocation?.allocate(4 * (intStr.length + Math.abs(e) + 2))
  if (e >= 0) return Number(intStr.padEnd(intStr.length + e, '0'))
  const cut = intStr.length + e
  if (cut >= 1) return Number(`${intStr.slice(0, cut)}.${intStr.slice(cut)}`)
  return Number(`0.${'0'.repeat(-cut)}${intStr}`)
}

/** digits of a non-negative rounded value at fixed decimals (expands e-notation). */
function digitsAt(rounded: number, decimals: number, allocation?: FormatAllocation): { intDigits: string; fracDigits: string } {
  allocation?.allocate(256) // finite repr, mantissa slices and exponent captures
  let s = String(rounded)
  if (s.startsWith('-')) s = s.slice(1)
  let intPart = s
  let fracPart = ''
  const eIdx = s.search(/[eE]/)
  if (eIdx >= 0) {
    const mant = s.slice(0, eIdx)
    const exp = Number(s.slice(eIdx + 1))
    const dot = mant.indexOf('.')
    const raw = mant.replace('.', '')
    const pointPos = (dot >= 0 ? dot : mant.length) + exp
    if (pointPos <= 0) {
      intPart = '0'
      allocation?.allocate(4 * (-pointPos + raw.length))
      fracPart = '0'.repeat(-pointPos) + raw
    } else if (pointPos >= raw.length) {
      allocation?.allocate(2 * pointPos)
      intPart = raw.padEnd(pointPos, '0')
      fracPart = ''
    } else {
      intPart = raw.slice(0, pointPos)
      fracPart = raw.slice(pointPos)
    }
  } else {
    const dot = s.indexOf('.')
    if (dot >= 0) { intPart = s.slice(0, dot); fracPart = s.slice(dot + 1) }
  }
  allocation?.allocate(4 * (fracPart.length + decimals) + 2 * decimals)
  fracPart = (fracPart + '0'.repeat(decimals)).slice(0, decimals)
  return { intDigits: intPart, fracDigits: fracPart }
}

function isDateLetter(ch: string): boolean {
  return ch === 'y' || ch === 'm' || ch === 'd' || ch === 'h' || ch === 's'
}

/** At most the four sections Excel reads (positive;negative;zero;text): the
 * split stops at the fourth unquoted ';' and ignores the rest, so its work is
 * bounded by the code prefix it keeps, not by the number of separators. With
 * authored codes capped at MAX_FORMAT_CODE_UNITS (parse.ts) a cell's format
 * parse is a per-cell constant; the allocation observer still prices TEXT()
 * format arguments, which are unbounded formula strings. */
function splitSections(format: string, allocation?: FormatAllocation): { sections: string[]; error?: string } {
  const sections: string[] = []
  let cur = ''
  let i = 0
  while (i < format.length) {
    const ch = format[i]
    if (ch === '"') {
      const end = format.indexOf('"', i + 1)
      if (end < 0) return { sections: [], error: 'unterminated quoted literal' }
      allocation?.allocate(2 * (end + 1 - i))
      cur = appendText(cur, format.slice(i, end + 1), allocation)
      i = end + 1
      continue
    }
    if (ch === '\\') {
      if (i + 1 >= format.length) return { sections: [], error: 'trailing escape' }
      allocation?.allocate(4)
      cur = appendText(cur, ch + format[i + 1], allocation)
      i += 2
      continue
    }
    if (ch === '[') {
      const end = format.indexOf(']', i + 1)
      if (end < 0) return { sections: [], error: 'unterminated bracket token' }
      allocation?.allocate(2 * (end + 1 - i))
      cur = appendText(cur, format.slice(i, end + 1), allocation)
      i = end + 1
      continue
    }
    if (ch === ';') {
      allocation?.allocate(8)
      sections.push(cur)
      if (sections.length === 4) return { sections }
      cur = ''
      i += 1
      continue
    }
    cur = appendText(cur, ch, allocation)
    i += 1
  }
  allocation?.allocate(8)
  sections.push(cur)
  return { sections }
}

function bracketReason(content: string): string {
  const c = content.toLowerCase()
  if (/^(red|green|blue|black|white|yellow|cyan|magenta|color\s?\d+)$/.test(c)) return 'color token'
  if (/^[=<>]/.test(c)) return 'condition token'
  if (c.startsWith('$-') || c.startsWith('$')) return 'currency/locale token'
  if (/^[hms]\b/.test(c) || /^[hms]\]/.test(c) || /^[hms]$/.test(c)) return 'elapsed token'
  return 'bracket token'
}

/** Tokenize one numeric/date section. Returns null on unsupported constructs. */
function tokenize(section: string, allocation?: FormatAllocation): { tokens: Tok[]; unsupported?: string } {
  const tokens: Tok[] = []
  let i = 0
  while (i < section.length) {
    allocation?.allocate(64) // token object and array slot, before any push
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      if (end < 0) return { tokens: [], unsupported: 'unterminated quoted literal' }
      allocation?.allocate(2 * (end - i - 1))
      tokens.push({ t: 'lit', text: section.slice(i + 1, end) })
      i = end + 1
      continue
    }
    if (ch === '\\') {
      if (i + 1 >= section.length) return { tokens: [], unsupported: 'trailing escape' }
      tokens.push({ t: 'lit', text: section[i + 1] })
      i += 2
      continue
    }
    if (ch === '[') {
      const end = section.indexOf(']', i + 1)
      if (end < 0) return { tokens: [], unsupported: 'unterminated bracket token' }
      allocation?.allocate(4 * (end - i - 1))
      return { tokens: [], unsupported: bracketReason(section.slice(i + 1, end)) }
    }
    if (ch === '0' || ch === '#') {
      tokens.push({ t: 'ph', ch })
      i += 1
      continue
    }
    if (ch === '?') return { tokens: [], unsupported: "'?' placeholder (not in the finite grammar)" }
    if (ch === '_') return { tokens: [], unsupported: "'_' skip-width token" }
    if (ch === '*') return { tokens: [], unsupported: "'*' fill token" }
    if (ch === '.') { tokens.push({ t: 'dot' }); i += 1; continue }
    if (ch === ',') { tokens.push({ t: 'comma' }); i += 1; continue }
    if (ch === '%') { tokens.push({ t: 'pct' }); i += 1; continue }
    if (ch === '@') { tokens.push({ t: 'at' }); i += 1; continue }
    if (ch === 'G') {
      allocation?.allocate(28)
      if (section.slice(i, i + 7).toLowerCase() === 'general') {
        return { tokens: [], unsupported: "'General' token" }
      }
      return { tokens: [], unsupported: `unexpected character 'G'` }
    }
    // Date tokens are recognized case-insensitively as an APPLICATION
    // INTEROPERABILITY PROFILE (MS-OI29500 §18.8.30 ABNF declares small
    // letters only; uppercase equivalence is not a claimed native rule). The
    // normalized token carries the small letter so downstream disambiguation
    // and rendering are identical to the lowercase form.
    const dateCh = ch.toLowerCase()
    if (isDateLetter(dateCh)) {
      let len = 0
      while (i + len < section.length && section[i + len].toLowerCase() === dateCh) len++
      tokens.push({ t: 'dt', ch: dateCh as Tok extends { t: 'dt'; ch: infer C } ? C : never, len })
      i += len
      continue
    }
    if ((ch === 'a' || ch === 'A' || ch === 'p' || ch === 'P')) {
      allocation?.allocate(20)
      const run = section.slice(i, i + 5)
      if (run.toLowerCase() === 'am/pm') {
        tokens.push({ t: 'am' })
        i += 5
        continue
      }
      return { tokens: [], unsupported: `unexpected character '${ch}'` }
    }
    if (ch === 'E' || ch === 'e') {
      return { tokens: [], unsupported: 'scientific token' }
    }
    // space and ASCII punctuation act as literal text in the finite grammar
    if (ch === ' ' || (ch.charCodeAt(0) >= 32 && ch.charCodeAt(0) < 127)) {
      tokens.push({ t: 'lit', text: ch })
      i += 1
      continue
    }
    return { tokens: [], unsupported: `unsupported character '${ch}'` }
  }
  return { tokens }
}

function hasDateTokens(tokens: Tok[]): boolean {
  return tokens.some((tk) => tk.t === 'dt' || tk.t === 'am')
}

function countDecimals(tokens: Tok[]): number {
  let count = 0
  let seenDot = false
  for (const tk of tokens) {
    if (tk.t === 'dot') { seenDot = true; continue }
    if (tk.t === 'ph') { if (seenDot) count++ }
  }
  return count
}

function analyzeNumeric(tokens: Tok[]): {
  scale: number; divisor: number; decimals: number; grouping: boolean
} {
  let pct = 0
  for (const tk of tokens) if (tk.t === 'pct') pct++
  const scale = Math.pow(100, pct) // percent multiplies (measured 0.125 -> '12.5%')
  // trailing scaling commas divide by 1000 each (measured 1234567 -> '1')
  let lastPh = -1
  tokens.forEach((tk, idx) => { if (tk.t === 'ph') lastPh = idx })
  let trailing = 0
  for (let idx = lastPh + 1; idx < tokens.length; idx++) {
    if (tokens[idx].t === 'comma') trailing++
    else if (tokens[idx].t === 'dot' || tokens[idx].t === 'ph') break
  }
  // grouping commas: commas BETWEEN digit placeholders (before the last ph)
  let grouping = false
  for (let idx = 0; idx < lastPh; idx++) {
    if (tokens[idx].t === 'comma') grouping = true
  }
  const decimals = countDecimals(tokens)
  return { scale, divisor: Math.pow(1000, trailing), decimals, grouping }
}

/**
 * Right-anchored digit distribution: the rightmost placeholder takes the last
 * digit; overflow digits attach to the LEFTMOST placeholder (after any prefix
 * literals). Literal-only sections (no placeholders) render no digits at all.
 *
 * SPEC r3 integer-# significance fix: a '0' digit PRESENT in the value is
 * significant (positional) and must never be stripped from '#' — only an
 * ABSENT digit ('' at idx<0) is omitted, and the leading zero of a
 * fractional-only value (<1, integer digit string "0") is suppressed in '#'.
 * '0' placeholders always emit (pad with '0', hold positions).
 */
function perPhChars(intDigits: string, intTokens: readonly Tok[], allocation?: FormatAllocation): string[] {
  const phs: Tok[] = []
  for (const tk of intTokens) if (tk.t === 'ph') { allocation?.allocate(8); phs.push(tk) }
  const n = intDigits.length
  const intIsFractionalZero = !/[1-9]/.test(intDigits) // e.g. value 0.47 -> "0"
  allocation?.allocate(phs.length * 8 + intDigits.length * 2)
  const chars = phs.map((ph, i) => {
    const idx = n - phs.length + i
    if (i === 0 && n > phs.length) return intDigits.slice(0, n - phs.length + 1)
    if (idx >= 0 && idx < n) {
      const digit = intDigits[idx]
      // optional '#' suppresses ONLY the leading zero of a fractional value
      if (
        ph.t === 'ph' && ph.ch === '#' && digit === '0' &&
        intIsFractionalZero
      ) return ''
      return digit
    }
    return ph.t === 'ph' && ph.ch === '0' ? '0' : ''
  })
  return chars
}

/**
 * Render the integer part walking tokens L->R with the right-anchored per-ph
 * digits. SPEC F7: when the format has grouping commas (between digit
 * placeholders), the COMPLETE emitted digit string (incl. overflow digits) is
 * grouped in 3s from the right — matching the measured single-comma patterns
 * ("#,##0": 12 -> "12"; 1234567 -> "1,234,567" with no leading comma) and
 * letting one comma position carry repeated separators. Commas after the
 * last placeholder remain scaling commas (emit nothing; divisor in
 * analyzeNumeric). Literal runs inside the field are preserved.
 */
function renderInteger(intDigits: string, intTokens: readonly Tok[], grouping: boolean, allocation?: FormatAllocation): string {
  const chars = perPhChars(intDigits, intTokens, allocation)
  let out = ''
  let phIdx = 0
  for (let i = 0; i < intTokens.length; i++) {
    const tk = intTokens[i]
    // Grouping needs the complete digit run (commas anchor at its right edge): the
    // output cap applies to the final emission, never to this intermediate.
    if (tk.t === 'lit') { out = appendText(out, tk.text, allocation); continue }
    if (tk.t === 'ph') { out = appendText(out, chars[phIdx++] ?? '', allocation); continue }
    if (tk.t === 'comma') { continue }
    if (tk.t === 'pct') { out = appendText(out, '%', allocation); continue }
    return '' // '@'/'dt'/'am' rejected by caller
  }
  if (grouping) {
    allocation?.allocate(4 * (out.length + Math.floor(out.length / 3)))
    out = out.replace(/\d+/g, (run) => run.replace(/\B(?=(\d{3})+(?!\d))/g, ','))
  }
  return out
}

function renderDateTime(
  serial: number, tokens: Tok[], semantics: ResolvedSemantics, allocation?: FormatAllocation,
): FormatResult {
  const system = semantics.dateSystem
  const abs = Math.abs(serial)
  const dayPart = Math.trunc(abs)
  const frac = abs - Math.trunc(abs) // raw fraction retained for time rendering
  const dec = decodeSerial(dayPart, system)
  if (isEvaluationError(dec)) return dec
  const civil = dec.civil
  if (semantics.locale !== 'en-US') {
    return { kind: 'unsupported', reason: 'date/time locale gate: only en-US profiles supported' }
  }
  const tod = timeOfDayParts(frac)
  const hasAm = tokens.some((tk) => tk.t === 'am')

  // month/minute disambiguation: m/mm adjacent to h (skipping literal
  // separators) or to s = minute
  allocation?.allocate(tokens.length * 8)
  const kinds: ('minute' | 'month' | 'other')[] = tokens.map((tk) =>
    tk.t === 'dt' && tk.ch === 'm' ? 'month' : 'other',
  )
  const prevSignificant = (i: number): Tok | undefined => {
    for (let k = i - 1; k >= 0; k--) if (tokens[k].t !== 'lit') return tokens[k]
    return undefined
  }
  const nextSignificant = (i: number): Tok | undefined => {
    for (let k = i + 1; k < tokens.length; k++) if (tokens[k].t !== 'lit') return tokens[k]
    return undefined
  }
  for (let i = 0; i < tokens.length; i++) {
    if (kinds[i] !== 'month') continue
    const tk = tokens[i] as Extract<Tok, { t: 'dt' }>
    if (tk.len > 2) continue // mmm/mmmm are month names
    const prevHour = (() => { const p = prevSignificant(i); return p && p.t === 'dt' && p.ch === 'h' })()
    const nextSec = (() => { const n = nextSignificant(i); return n && n.t === 'dt' && n.ch === 's' })()
    if (prevHour || nextSec) kinds[i] = 'minute'
  }

  let out = ''
  for (let i = 0; i < tokens.length; i++) {
    const tk = tokens[i]
    allocation?.allocate(64) // bounded date digits/padding and names
    if (tk.t === 'lit') { out = appendOut(out, tk.text, allocation); continue }
    if (tk.t === 'comma') { out = appendOut(out, ',', allocation); continue } // literal separator in date sections
    if (tk.t === 'dot') { out = appendOut(out, '.', allocation); continue }
    if (tk.t === 'am') {
      const h12 = ((tod.hour + 11) % 12) + 1
      out = appendOut(out, (tod.hour >= 12 ? 'PM' : 'AM'), allocation)
      void h12
      continue
    }
    if (tk.t !== 'dt') return { kind: 'unsupported', reason: 'non-date token in a date/time section' }
    switch (tk.ch) {
      case 'y': {
        const y = String(civil.year)
        out = appendOut(out, tk.len >= 3 ? y : y.slice(-2).padStart(2, '0'), allocation)
        break
      }
      case 'm': {
        if (kinds[i] === 'minute') {
          out = appendOut(out, tk.len >= 2 ? String(tod.minute).padStart(2, '0') : String(tod.minute), allocation)
        } else if (tk.len >= 4) {
          out = appendOut(out, MONTHS[civil.month - 1], allocation)
        } else if (tk.len === 3) {
          out = appendOut(out, MONTHS_SHORT[civil.month - 1], allocation)
        } else {
          out = appendOut(out, tk.len === 2 ? String(civil.month).padStart(2, '0') : String(civil.month), allocation)
        }
        break
      }
      case 'd': {
        if (tk.len >= 3) {
          const w = weekdayIndex(dayPart, system)
          if (isEvaluationError(w)) return w
          out = appendOut(out, tk.len >= 4 ? WEEKDAYS[w] : WEEKDAYS_SHORT[w], allocation)
        } else {
          out = appendOut(out, tk.len === 2 ? String(civil.day).padStart(2, '0') : String(civil.day), allocation)
        }
        break
      }
      case 'h': {
        const h = hasAm ? ((tod.hour + 11) % 12) + 1 : tod.hour
        out = appendOut(out, tk.len >= 2 ? String(h).padStart(2, '0') : String(h), allocation)
        break
      }
      case 's': {
        out = appendOut(out, tk.len >= 2 ? String(tod.second).padStart(2, '0') : String(tod.second), allocation)
        break
      }
    }
  }
  return { kind: 'text', text: capOut(out, allocation) }
}

/**
 * True when a NUMERIC value is shown through a text-format section ('@'),
 * using formatCommon's own section selection (positive;negative;zero; the
 * fourth, text section never applies to numbers). Quoted/escaped literals are
 * tokens of their own, so an '@' inside quotes does not count. The renderer
 * uses this to keep '#' overflow for numeric sections only.
 */
export function numericSectionIsText(value: number, format: string, allocation?: FormatAllocation): boolean {
  // Paint calls this after formatting: its builders are admitted by the same
  // observer, and the split keeps at most four sections.
  const { sections, error } = splitSections(format, allocation)
  if (error) return false
  let sectionIdx = 0
  if (value < 0 && sections.length >= 2) sectionIdx = 1
  else if (value === 0 && sections.length >= 3) sectionIdx = 2
  const parsed = tokenize(sections[sectionIdx] ?? '', allocation)
  return !parsed.unsupported && parsed.tokens.some(tk => tk.t === 'at')
}

/** Whether one already-selected section contains supported date/time tokens. */
export function dateSection(format: string, allocation?: FormatAllocation): boolean {
  const parsed = tokenize(format, allocation)
  return !parsed.unsupported && hasDateTokens(parsed.tokens)
}

export function formatCommon(
  value: ElementValue,
  format: string,
  semantics: ResolvedSemantics,
  allocation?: FormatAllocation,
): FormatResult {
  if (isEvaluationError(value)) return value // tagged errors propagate
  const { sections: secList, error: splitErr } = splitSections(format, allocation)
  if (splitErr) return { kind: 'unsupported', reason: splitErr }

  // SPEC F8: text/boolean/null values take the FOURTH (text) section when it
  // exists: quoted/escaped literals render, '@' inserts the value text, and an
  // empty section or one WITHOUT '@' hides the text. With no fourth section
  // the value passes through unchanged (unchanged r1/r2 behavior for 1-3
  // section formats; a bare '@' section is the .0#-scoped literal transport).
  if (typeof value === 'string' || typeof value === 'boolean' || value === null) {
    const textSec = secList.length >= 4 ? secList[3] : null
    if (textSec !== null) {
      if (textSec === '') return { kind: 'text', text: '' } // empty text section hides
      const t = tokenize(textSec, allocation)
      if (t.unsupported) return { kind: 'unsupported', reason: t.unsupported }
      let outText = ''
      let sawAt = false
      for (const tk of t.tokens) {
        if (tk.t === 'lit') { outText = appendOut(outText, tk.text, allocation); continue }
        if (tk.t === 'at') {
          sawAt = true
          outText = appendOut(outText, typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value ?? '', allocation)
          continue
        }
        if (tk.t === 'ph' || tk.t === 'comma' || tk.t === 'dot' || tk.t === 'pct') {
          continue // omitted-@; digits hold no meaning in the text section here
        }
        return { kind: 'unsupported', reason: 'date token in the text section' }
      }
      void sawAt
      return { kind: 'text', text: capOut(outText, allocation) }
    }
    const asText = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value ?? ''
    if (typeof value === 'string') allocation?.borrowed?.(value)
    return { kind: 'text', text: capOut(asText, allocation) }
  }

  const v = value
  if (!Number.isFinite(v)) return { kind: 'unsupported', reason: 'non-finite numeric value' }

  // section selection
  let sectionIdx = 0
  let magnitude = Math.abs(v)
  let negativePrefix = ''
  if (v < 0) {
    if (secList.length >= 2) { sectionIdx = 1; magnitude = Math.abs(v) }
    else negativePrefix = '-'
  } else if (v === 0 && secList.length >= 3) {
    sectionIdx = 2
  }
  const section = secList[sectionIdx]
  const parsed = tokenize(section, allocation)
  if (parsed.unsupported) return { kind: 'unsupported', reason: parsed.unsupported }
  const tokens = parsed.tokens

  if (hasDateTokens(tokens)) {
    if (v < 0) {
      // measured 1904 row: negative serials format as '-' + format(|serial|)
      const r = formatCommon(magnitude, format, semantics, allocation)
      if (r.kind === 'text') { allocation?.allocate(2 * (r.text.length + 1)); return { kind: 'text', text: capOut('-' + r.text, allocation) } }
      return r
    }
    return renderDateTime(v, tokens, semantics, allocation)
  }

  // numeric rendering: percent multiplies; trailing scaling commas divide
  const { scale, divisor, decimals, grouping } = analyzeNumeric(tokens)
  const scaled = (magnitude * scale) / divisor
  // QUALITY Q2 fix: internal overflow/NaN is NEVER labeled computed text. The
  // exact Excel result at such numeric boundaries is unmeasured; the formatter
  // policy (capability boundary) returns a diagnosed unsupported outcome.
  if (!Number.isFinite(scaled)) {
    return {
      kind: 'unsupported',
      reason: 'numeric-limit: percent/scaling arithmetic exceeded the finite formatter range for a finite input (exact Excel result unmeasured at this boundary; policy: diagnosed unsupported, cache retention is owner-side)',
    }
  }
  const rounded = roundHalfAwayDecimal(scaled, decimals, allocation)
  if (!Number.isFinite(rounded)) {
    return {
      kind: 'unsupported',
      reason: 'numeric-limit: rounded value exceeded the finite formatter range (exact Excel result unmeasured at this boundary; policy: diagnosed unsupported, cache retention is owner-side)',
    }
  }
  const { intDigits, fracDigits } = digitsAt(rounded, decimals, allocation)
  const dotIdx = tokens.findIndex((tk) => tk.t === 'dot')
  allocation?.allocate(tokens.length * 8)
  const intTokens = dotIdx >= 0 ? tokens.slice(0, dotIdx) : tokens
  const fracTokens = dotIdx >= 0 ? tokens.slice(dotIdx + 1) : []
  const intStr = renderInteger(intDigits, intTokens, grouping, allocation)

  allocation?.allocate(2 * (negativePrefix.length + intStr.length))
  let out = negativePrefix + intStr
  if (dotIdx >= 0) {
    out = appendOut(out, '.', allocation)
    // SPEC F6: fraction '#/'0' emission with the same optional-# trailing rule.
    // QUALITY Q3 fix: suffix significance is computed in ONE right-to-left
    // pass (linear) instead of scanning the remaining suffix per placeholder
    // (which was O(n^2) in the number of optional fraction placeholders).
    allocation?.allocate(fracTokens.length * 16)
    const fracPhs = fracTokens.filter((tk) => tk.t === 'ph') as Extract<Tok, { t: 'ph' }>[]
    let placed = 0
    const fracChars: string[] = new Array(fracPhs.length)
    let nonzeroRight = false
    for (let i = fracPhs.length - 1; i >= 0; i--) {
      const d = fracDigits[i] ?? '0'
      if (fracPhs[i].ch === '0') {
        fracChars[i] = d // required placeholder always emits
      } else if (d !== '0') {
        fracChars[i] = d
      } else {
        fracChars[i] = nonzeroRight ? d : '' // omitted trailing zero in '#'
      }
      if (d !== '0') nonzeroRight = true
    }
    for (const tk of fracTokens) {
      if (tk.t === 'ph') {
        out = appendOut(out, fracChars[placed] ?? '', allocation)
        placed++
        continue
      }
      if (tk.t === 'lit') { out = appendOut(out, tk.text, allocation); continue }
      if (tk.t === 'dot') { out = appendOut(out, '.', allocation); continue }
      if (tk.t === 'comma') { continue } // scaling comma emits nothing
      if (tk.t === 'pct') { out = appendOut(out, '%', allocation); continue }
      return { kind: 'unsupported', reason: 'unexpected token after decimal point' }
    }
  } else {
    for (const tk of tokens) {
      if (tk.t === 'at') return { kind: 'unsupported', reason: "'@' in a numeric section" }
      if (tk.t === 'dt' || tk.t === 'am') {
        return { kind: 'unsupported', reason: 'date token in a numeric section' }
      }
    }
  }
  return { kind: 'text', text: capOut(out, allocation) }
}

export { isEvaluationError }
