/**
 * Cell-display number formats beyond the finite TEXT grammar (visual review #7).
 *
 * `formula/format.ts` stays the measured, finite grammar shared with TEXT and
 * diagnoses colours, scientific notation, fractions and padding tokens. Excel
 * displays those in cells, so the renderer (only) layers them here:
 *
 * - colour sections ([Red], [Blue], …, [Color n]) choose the painted text colour;
 * - `_x` reserves a space, `*x` fill is dropped, `?` pads like `#`;
 * - [$-409] locale tags drop and [$€-407] currency tags become their literal;
 * - scientific mantissa/exponent (`0.00E+00`, engineering `##0.0E+0`);
 * - fractions (`# ?/?`, `# ??/??`, fixed denominators such as `# ?/8`);
 * - a `General` section.
 *
 * Conditions ([>100]) and elapsed time ([h]) are still unsupported; the caller
 * keeps its diagnostic and fallback for anything this returns undefined for.
 * Only this lazily loaded renderer path imports the module.
 */
import { formatCommon, type FormatAllocation } from './formula/format'
import type { ElementValue, ResolvedSemantics } from './formula/types'

const NAMED: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
  magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
}
/** Excel's default 56-colour palette, `[Color1]`..`[Color56]` (indexed 8..63). */
const PALETTE = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 '
  + '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF '
  + '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 '
  + '003366 339966 003300 333300 993300 993366 333399 333333').split(' ')

/** Built-in numFmtIds whose codes need this layer (ECMA-376 §18.8.30). */
export const BUILTIN_EXTENDED_FORMATS: Readonly<Record<number, string>> = {
  1: '0', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  48: '##0.0E+0',
}

interface Section { body: string; color?: string; condition: boolean; unsupported: boolean }

/** Split on unquoted, unescaped ';' and strip each section's bracket tokens. */
function sections(code: string): Section[] {
  const out: Section[] = []
  let body = '', color: string | undefined, condition = false, unsupported = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      if (end < 0) { unsupported = true; break }
      body += code.slice(i, end + 1); i = end
    } else if (ch === '\\') {
      body += code.slice(i, i + 2); i++
    } else if (ch === '[') {
      const end = code.indexOf(']', i + 1)
      if (end < 0) { unsupported = true; break }
      const tag = code.slice(i + 1, end), lower = tag.toLowerCase()
      const indexed = /^color\s?(\d+)$/.exec(lower)
      if (NAMED[lower]) color = NAMED[lower]
      else if (indexed) { const n = Number(indexed[1]); if (n >= 1 && n <= 56) color = `#${PALETTE[n - 1]}`; else unsupported = true }
      else if (/^[<>=]/.test(lower)) condition = true
      else if (lower.startsWith('$')) {
        // [$€-407]: the currency literal; [$-409]: a locale tag only.
        const symbol = tag.slice(1).split('-')[0]
        if (symbol) body += `"${symbol.replace(/"/g, '')}"`
      } else unsupported = true
      i = end
    } else if (ch === ';') {
      out.push({ body, color, condition, unsupported })
      body = ''; color = undefined; condition = false; unsupported = false
    } else if (ch === '_') {
      if (i + 1 < code.length) { body += '" "'; i++ }
    } else if (ch === '*') {
      i++ // repeat-fill to the cell width: dropped
    } else body += ch
  }
  out.push({ body, color, condition, unsupported })
  return out.slice(0, 4)
}

/** Excel section selection for a value (no conditions). */
function select(value: ElementValue, list: Section[]): { index: number; signed: boolean } {
  if (typeof value !== 'number') return { index: list.length >= 4 ? 3 : -1, signed: false }
  if (value < 0 && list.length >= 2) return { index: 1, signed: false }
  if (value === 0 && list.length >= 3) return { index: 2, signed: false }
  return { index: 0, signed: value < 0 }
}

/** Mask quoted/escaped literals so structural scans ignore their characters. */
function structural(body: string): string {
  return body.replace(/"[^"]*"/g, m => '\u0001'.repeat(m.length)).replace(/\\./g, '\u0001\u0001')
}

/** The painted colour a format's selected section asks for, if any. */
export function formatColor(value: ElementValue, code: string): string | undefined {
  if (!code.includes('[')) return undefined
  const list = sections(code)
  if (list.some(s => s.condition || s.unsupported)) return undefined
  const { index } = select(value, list)
  return index >= 0 ? list[index]?.color : undefined
}

/**
 * Cell text for formats the finite grammar diagnoses; undefined when this layer
 * cannot render the code either (the caller then keeps its fallback).
 */
export function formatExtended(value: ElementValue, code: string, semantics: ResolvedSemantics, general: (n: number) => string, allocation?: FormatAllocation): string | undefined {
  const list = sections(code)
  if (list.some(s => s.condition || s.unsupported)) return undefined
  allocation?.allocate(4 * code.length + 512) // bounded rebuilt sections and output
  const { index, signed } = select(value, list)
  if (typeof value === 'number' && Number.isFinite(value) && index >= 0) {
    const body = list[index].body, shape = structural(body)
    const magnitude = Math.abs(value), sign = signed ? '-' : ''
    if (/^\s*general\s*$/i.test(shape)) return sign + general(magnitude)
    if (/E[+-]/i.test(shape)) { const t = scientific(magnitude, body, shape, semantics, allocation); return t === undefined ? undefined : sign + t }
    if (shape.includes('/') && !/[ymdhs]/i.test(shape.replace(/E/gi, ''))) { const t = fraction(magnitude, shape); return t === undefined ? undefined : sign + t }
    // Plain numeric/date body: '?' pads like '#'. The selected section alone
    // goes to the finite grammar (its sign already decided above).
    const result = formatCommon(signed ? value : magnitude, body.replace(/\?/g, (m, at: number) => shape[at] === '?' ? '#' : m), semantics, allocation)
    return result.kind === 'text' ? result.text : undefined
  }
  // Text values (and numbers whose section is absent) use the cleaned code.
  const cleaned = list.map(s => s.body.replace(/\?/g, '#')).join(';')
  const result = formatCommon(value, cleaned, semantics, allocation)
  return result.kind === 'text' ? result.text : undefined
}

/** `0.00E+00`, `##0.0E+0`: the mantissa through the finite grammar, an exponent
 * that is a multiple of the integer placeholder count (engineering notation). */
function scientific(x: number, body: string, shape: string, semantics: ResolvedSemantics, allocation?: FormatAllocation): string | undefined {
  const at = shape.search(/E[+-]/i)
  const mantissaCode = body.slice(0, at), rest = body.slice(at + 2), plus = shape[at + 1] === '+'
  const digits = /^[0#?]+/.exec(rest)?.[0] ?? ''
  if (!digits) return undefined
  const intPart = structural(mantissaCode).split('.')[0]
  const step = Math.max(1, (intPart.match(/[0#?]/g) ?? []).length)
  const decimals = (structural(mantissaCode).split('.')[1]?.match(/[0#?]/g) ?? []).length
  let exponent = x === 0 ? 0 : Math.floor(Math.log10(x))
  exponent = step > 1 ? Math.floor(exponent / step) * step : exponent
  let mantissa = x / 10 ** exponent
  // Rounding can carry the mantissa to the next power (9.999 -> 10.00).
  if (x !== 0 && Number(mantissa.toFixed(decimals)) >= 10 ** step) { exponent += step; mantissa = x / 10 ** exponent }
  const head = formatCommon(mantissa, mantissaCode.replace(/\?/g, '#'), semantics, allocation)
  if (head.kind !== 'text') return undefined
  const e = String(Math.abs(exponent)).padStart((digits.match(/0/g) ?? []).length, '0')
  const tail = rest.slice(digits.length).replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1')
  return `${head.text}E${exponent < 0 ? '-' : plus ? '+' : ''}${e}${tail}`
}

/** `# ?/?`, `# ??/??`, `?/?`, `# ?/8`: an optional whole part, then the closest
 * fraction whose denominator fits the placeholders (or the fixed one). */
function fraction(x: number, shape: string): string | undefined {
  const m = /^\s*(?:([#0?,]+)\s+)?([#0?]+)\s*\/\s*([#0?]+|\d+)\s*$/.exec(shape)
  if (!m) return undefined
  const [, wholePh, numPh, denPh] = m
  let whole = wholePh ? Math.floor(x) : 0
  const part = wholePh ? x - whole : x
  let num: number, den: number
  if (/^\d+$/.test(denPh)) {
    den = Number(denPh)
    if (!(den > 0)) return undefined
    num = Math.round(part * den)
  } else {
    const max = 10 ** denPh.length - 1
    num = Math.round(part); den = 1
    let best = Math.abs(part - num)
    for (let d = 2; d <= max && best > 0; d++) {
      const n = Math.round(part * d), error = Math.abs(part - n / d)
      if (error < best - 1e-12) { best = error; num = n; den = d }
    }
  }
  if (wholePh && num >= den) { whole += Math.floor(num / den); num %= den }
  const width = numPh.length + 1 + denPh.length
  const wholeText = whole > 0 ? String(whole) : wholePh?.includes('0') ? '0' : ''
  if (num === 0) return wholePh ? (wholeText || '0') + ' '.repeat(width + 1) : `0/${denPh.replace(/[#0?]/g, '1').slice(0, 1)}`
  const numText = String(num).padStart(numPh.length, numPh.includes('0') ? '0' : ' ')
  const denText = /^\d+$/.test(denPh) ? denPh : String(den).padEnd(denPh.length, ' ')
  return `${wholeText ? `${wholeText} ` : wholePh ? ' ' : ''}${numText}/${denText}`
}
