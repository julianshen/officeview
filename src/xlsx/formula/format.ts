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
export function roundHalfAwayDecimal(x: number, decimals: number): number {
  if (!Number.isFinite(x)) return x
  if (x === 0) return 0
  const neg = x < 0
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
  let kept = keep >= digits.length ? digits.padEnd(keep, '0') : digits.slice(0, keep)
  if (up) {
    // increment the kept digit string (BigInt for safety)
    kept = (BigInt(kept) + 1n).toString()
  }
  // value = int(kept) * 10^(point - keep) = int(kept) * 10^(-decimals)
  const mag = stringTimesPow10(kept, point - keep)
  return neg ? -mag : mag
}

/** exact decimal-string times 10^e, converted to Number (lossless for modest sizes). */
function stringTimesPow10(intStr: string, e: number): number {
  if (e >= 0) return Number(intStr.padEnd(intStr.length + e, '0'))
  const cut = intStr.length + e
  if (cut >= 1) return Number(`${intStr.slice(0, cut)}.${intStr.slice(cut)}`)
  return Number(`0.${'0'.repeat(-cut)}${intStr}`)
}

/** digits of a non-negative rounded value at fixed decimals (expands e-notation). */
function digitsAt(rounded: number, decimals: number): { intDigits: string; fracDigits: string } {
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
      fracPart = '0'.repeat(-pointPos) + raw
    } else if (pointPos >= raw.length) {
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
  fracPart = (fracPart + '0'.repeat(decimals)).slice(0, decimals)
  return { intDigits: intPart, fracDigits: fracPart }
}

function isDateLetter(ch: string): boolean {
  return ch === 'y' || ch === 'm' || ch === 'd' || ch === 'h' || ch === 's'
}

function splitSections(format: string): { sections: string[]; error?: string } {
  const sections: string[] = []
  let cur = ''
  let i = 0
  while (i < format.length) {
    const ch = format[i]
    if (ch === '"') {
      const end = format.indexOf('"', i + 1)
      if (end < 0) return { sections: [], error: 'unterminated quoted literal' }
      cur += format.slice(i, end + 1)
      i = end + 1
      continue
    }
    if (ch === '\\') {
      if (i + 1 >= format.length) return { sections: [], error: 'trailing escape' }
      cur += ch + format[i + 1]
      i += 2
      continue
    }
    if (ch === '[') {
      const end = format.indexOf(']', i + 1)
      if (end < 0) return { sections: [], error: 'unterminated bracket token' }
      cur += format.slice(i, end + 1)
      i = end + 1
      continue
    }
    if (ch === ';') {
      sections.push(cur)
      cur = ''
      i += 1
      continue
    }
    cur += ch
    i += 1
  }
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
function tokenize(section: string): { tokens: Tok[]; unsupported?: string } {
  const tokens: Tok[] = []
  let i = 0
  while (i < section.length) {
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      if (end < 0) return { tokens: [], unsupported: 'unterminated quoted literal' }
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
function perPhChars(intDigits: string, intTokens: readonly Tok[]): string[] {
  const phs: Tok[] = []
  for (const tk of intTokens) if (tk.t === 'ph') phs.push(tk)
  const n = intDigits.length
  const intIsFractionalZero = !/[1-9]/.test(intDigits) // e.g. value 0.47 -> "0"
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
function renderInteger(intDigits: string, intTokens: readonly Tok[], grouping: boolean): string {
  const chars = perPhChars(intDigits, intTokens)
  let out = ''
  let phIdx = 0
  for (let i = 0; i < intTokens.length; i++) {
    const tk = intTokens[i]
    if (tk.t === 'lit') { out += tk.text; continue }
    if (tk.t === 'ph') { out += chars[phIdx++] ?? ''; continue }
    if (tk.t === 'comma') { continue }
    if (tk.t === 'pct') { out += '%'; continue }
    return '' // '@'/'dt'/'am' rejected by caller
  }
  if (grouping) {
    out = out.replace(/\d+/g, (run) => run.replace(/\B(?=(\d{3})+(?!\d))/g, ','))
  }
  return out
}

function renderDateTime(
  serial: number, tokens: Tok[], semantics: ResolvedSemantics,
): FormatResult {
  const system = semantics.dateSystem
  const abs = Math.abs(serial)
  const dayPart = Math.trunc(abs)
  const frac = abs - Math.trunc(abs) // raw fraction retained for time rendering
  const dec = decodeSerial(dayPart, system)
  if (isEvaluationError(dec)) return dec
  const civil = dec.civil
  if (semantics.locale !== 'en-US') {
    return { kind: 'unsupported', reason: 'locale names (month/weekday) gate: only en-US names supported' }
  }
  const tod = timeOfDayParts(frac)
  const hasAm = tokens.some((tk) => tk.t === 'am')

  // month/minute disambiguation: m/mm adjacent to h (skipping literal
  // separators) or to s = minute
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
    if (tk.t === 'lit') { out += tk.text; continue }
    if (tk.t === 'comma') { out += ','; continue } // literal separator in date sections
    if (tk.t === 'dot') { out += '.'; continue }
    if (tk.t === 'am') {
      const h12 = ((tod.hour + 11) % 12) + 1
      out += (tod.hour >= 12 ? 'PM' : 'AM')
      void h12
      continue
    }
    if (tk.t !== 'dt') return { kind: 'unsupported', reason: 'non-date token in a date/time section' }
    switch (tk.ch) {
      case 'y': {
        const y = String(civil.year)
        out += tk.len >= 3 ? y : y.slice(-2).padStart(2, '0')
        break
      }
      case 'm': {
        if (kinds[i] === 'minute') {
          out += tk.len >= 2 ? String(tod.minute).padStart(2, '0') : String(tod.minute)
        } else if (tk.len >= 4) {
          out += MONTHS[civil.month - 1]
        } else if (tk.len === 3) {
          out += MONTHS_SHORT[civil.month - 1]
        } else {
          out += tk.len === 2 ? String(civil.month).padStart(2, '0') : String(civil.month)
        }
        break
      }
      case 'd': {
        if (tk.len >= 3) {
          const w = weekdayIndex(dayPart, system)
          if (isEvaluationError(w)) return w
          out += tk.len >= 4 ? WEEKDAYS[w] : WEEKDAYS_SHORT[w]
        } else {
          out += tk.len === 2 ? String(civil.day).padStart(2, '0') : String(civil.day)
        }
        break
      }
      case 'h': {
        const h = hasAm ? ((tod.hour + 11) % 12) + 1 : tod.hour
        out += tk.len >= 2 ? String(h).padStart(2, '0') : String(h)
        break
      }
      case 's': {
        out += tk.len >= 2 ? String(tod.second).padStart(2, '0') : String(tod.second)
        break
      }
    }
  }
  return { kind: 'text', text: out }
}

export function formatCommon(
  value: ElementValue,
  format: string,
  semantics: ResolvedSemantics,
): FormatResult {
  if (isEvaluationError(value)) return value // tagged errors propagate
  const { sections, error: splitErr } = splitSections(format)
  if (splitErr) return { kind: 'unsupported', reason: splitErr }
  const secList = sections.slice(0, 4)

  // SPEC F8: text/boolean/null values take the FOURTH (text) section when it
  // exists: quoted/escaped literals render, '@' inserts the value text, and an
  // empty section or one WITHOUT '@' hides the text. With no fourth section
  // the value passes through unchanged (unchanged r1/r2 behavior for 1-3
  // section formats; a bare '@' section is the .0#-scoped literal transport).
  if (typeof value === 'string' || typeof value === 'boolean' || value === null) {
    const textSec = secList.length >= 4 ? secList[3] : null
    if (textSec !== null) {
      if (textSec === '') return { kind: 'text', text: '' } // empty text section hides
      const t = tokenize(textSec)
      if (t.unsupported) return { kind: 'unsupported', reason: t.unsupported }
      let outText = ''
      let sawAt = false
      for (const tk of t.tokens) {
        if (tk.t === 'lit') { outText += tk.text; continue }
        if (tk.t === 'at') {
          sawAt = true
          outText += typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value ?? ''
          continue
        }
        if (tk.t === 'ph' || tk.t === 'comma' || tk.t === 'dot' || tk.t === 'pct') {
          continue // omitted-@; digits hold no meaning in the text section here
        }
        return { kind: 'unsupported', reason: 'date token in the text section' }
      }
      void sawAt
      return { kind: 'text', text: outText }
    }
    const asText = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value ?? ''
    return { kind: 'text', text: asText }
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
  const parsed = tokenize(section)
  if (parsed.unsupported) return { kind: 'unsupported', reason: parsed.unsupported }
  const tokens = parsed.tokens

  if (hasDateTokens(tokens)) {
    if (v < 0) {
      // measured 1904 row: negative serials format as '-' + format(|serial|)
      const r = formatCommon(magnitude, format, semantics)
      if (r.kind === 'text') return { kind: 'text', text: '-' + r.text }
      return r
    }
    return renderDateTime(v, tokens, semantics)
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
  const rounded = roundHalfAwayDecimal(scaled, decimals)
  if (!Number.isFinite(rounded)) {
    return {
      kind: 'unsupported',
      reason: 'numeric-limit: rounded value exceeded the finite formatter range (exact Excel result unmeasured at this boundary; policy: diagnosed unsupported, cache retention is owner-side)',
    }
  }
  const { intDigits, fracDigits } = digitsAt(rounded, decimals)
  const dotIdx = tokens.findIndex((tk) => tk.t === 'dot')
  const intTokens = dotIdx >= 0 ? tokens.slice(0, dotIdx) : tokens
  const fracTokens = dotIdx >= 0 ? tokens.slice(dotIdx + 1) : []
  const intStr = renderInteger(intDigits, intTokens, grouping)

  let out = negativePrefix + intStr
  if (dotIdx >= 0) {
    out += '.'
    // SPEC F6: fraction '#/'0' emission with the same optional-# trailing rule.
    // QUALITY Q3 fix: suffix significance is computed in ONE right-to-left
    // pass (linear) instead of scanning the remaining suffix per placeholder
    // (which was O(n^2) in the number of optional fraction placeholders).
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
        out += fracChars[placed] ?? ''
        placed++
        continue
      }
      if (tk.t === 'lit') { out += tk.text; continue }
      if (tk.t === 'dot') { out += '.'; continue }
      if (tk.t === 'comma') { continue } // scaling comma emits nothing
      if (tk.t === 'pct') { out += '%'; continue }
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
  return { kind: 'text', text: out }
}

export { isEvaluationError }
