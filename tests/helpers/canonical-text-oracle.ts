/** Slow, test-only oracle written directly from spec §7.1 (amended 2026-10-11).
 * It deliberately shares no matching or segmentation code with production:
 * - canonical form: NFD, remove \p{Default_Ignorable_Code_Point} (host regex)
 *   except ZWJ, variation selectors and tag characters, then full case folding
 *   from a plain CaseFolding C+F listing, with the Turkic I/U+0307 conditions
 *   taken from SpecialCasing.txt and evaluated by direct scans over exact
 *   Canonical_Combining_Class values from UnicodeData.txt; then NFD;
 * - graphemes: host Intl.Segmenter (allowed here as a test-only cross-check),
 *   with every lone surrogate forced to be its own unit;
 * - matching: exhaustive enumeration of grapheme-boundary candidates comparing
 *   canonical strings, and memoized recursive glob over grapheme units.
 * The host regex/segmenter follow the host's Unicode version, so pools must use
 * characters whose properties are identical in the host and pinned versions.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { TextOp } from '../../src/xlsx/view/contracts'

const fixture = (name: string) => JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'unicode', name), 'utf8'))
const folding = new Map<number, string>()
for (const entry of (fixture('case-folding-16.0.0.json').full as string).split(',')) {
  const [point, target] = entry.split(':')
  folding.set(parseInt(point, 36), String.fromCodePoint(...target.split('.').map(value => parseInt(value, 36))))
}
const combiningClass = new Map<number, number>()
for (const entry of (fixture('combining-class-16.0.0.json').nonzero as string).split(',')) {
  const [point, ccc] = entry.split(':')
  combiningClass.set(parseInt(point, 36), Number(ccc))
}
const ccc = (point: string) => combiningClass.get(point.codePointAt(0)!) ?? 0

// Intl.Segmenter is ES2022; the project lib targets ES2020.
interface GraphemeSegmenter { segment(text: string): Iterable<{ segment: string }> }
const { Segmenter } = Intl as unknown as { Segmenter: new (locale: string, options: { granularity: 'grapheme' }) => GraphemeSegmenter }
const segmenter = new Segmenter('en', { granularity: 'grapheme' })
const lone = (point: string) => point.length === 1 && point.charCodeAt(0) >= 0xd800 && point.charCodeAt(0) <= 0xdfff
/** SpecialCasing.txt: in After_I and Before_Dot, characters of combining
 * class 0 or 230 (Above) interrupt the context; any other class may intervene. */
const interrupts = (point: string) => ccc(point) === 0 || ccc(point) === 230
/** After_I: there is an uppercase I before C, with no intervening class 0 or 230. */
function afterI(points: readonly string[], at: number): boolean {
  for (let j = at - 1; j >= 0; j--) {
    if (points[j] === 'I') return true
    if (interrupts(points[j])) return false
  }
  return false
}
/** Before_Dot: C is followed by U+0307, with no intervening class 0 or 230. */
function beforeDot(points: readonly string[], at: number): boolean {
  for (let j = at + 1; j < points.length; j++) {
    if (points[j] === '\u0307') return true
    if (interrupts(points[j])) return false
  }
  return false
}

export function oracleCanonical(text: string, locale: string): string {
  const turkic = ['tr', 'az'].includes(locale.split(/[-_]/)[0].toLowerCase())
  const points = Array.from(text.normalize('NFD').replace(/(?![\u200d\ufe00-\ufe0f\u{e0020}-\u{e007f}\u{e0100}-\u{e01ef}])\p{Default_Ignorable_Code_Point}/gu, ''))
  let folded = ''
  for (let i = 0; i < points.length; i++) {
    const point = points[i]
    // Turkic: I lowercases to dotless ı unless Before_Dot; U+0307 is removed
    // After_I (together these fold decomposed İ to i); İ itself folds to i (T).
    if (turkic && point === 'I') folded += beforeDot(points, i) ? 'i' : '\u0131'
    else if (turkic && point === '\u0307') folded += afterI(points, i) ? '' : point
    else if (turkic && point === '\u0130') folded += 'i'
    else folded += folding.get(point.codePointAt(0)!) ?? point
  }
  return folded.normalize('NFD')
}

export function oracleGraphemes(canonical: string): string[] {
  const result: string[] = []
  let piece = ''
  const flush = () => { for (const { segment } of segmenter.segment(piece)) result.push(segment); piece = '' }
  for (const point of canonical) {
    if (lone(point)) { flush(); result.push(point) } else piece += point
  }
  flush()
  return result
}

type Token = { kind: 'literal'; text: string } | { kind: 'star' } | { kind: 'one' }
function oracleTokens(pattern: string): Token[] {
  const tokens: Token[] = []
  let literal: string | undefined, escaped = false
  for (const point of pattern) {
    if (escaped) {
      if (!'*?~'.includes(point)) throw new RangeError('Wildcard escape must precede a metacharacter')
      literal = (literal ?? '') + point; escaped = false
    } else if (point === '~') escaped = true
    else if (point === '*' || point === '?') {
      if (literal !== undefined) tokens.push({ kind: 'literal', text: literal })
      literal = undefined
      tokens.push({ kind: point === '*' ? 'star' : 'one' })
    } else literal = (literal ?? '') + point
  }
  if (escaped) throw new RangeError('Malformed wildcard trailing escape')
  if (literal !== undefined) tokens.push({ kind: 'literal', text: literal })
  return tokens
}

export function oracleMatch(op: TextOp, pattern: string, wildcards: boolean, text: string, locale: string): boolean {
  const negative = op === 'neq' || op === 'not-contains'
  // A query whose canonical form is empty is no search term: it does not filter.
  const empty = wildcards ? oracleTokens(pattern).every(token => token.kind === 'literal' && !oracleCanonical(token.text, locale)) : !oracleCanonical(pattern, locale)
  if (empty) return true
  const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
  const canonical = oracleCanonical(text, locale)
  const units = oracleGraphemes(canonical)
  const at = [0]
  for (const unit of units) at.push(at[at.length - 1] + unit.length)
  const n = units.length
  const slice = (i: number, j: number) => canonical.slice(at[i], at[j])
  let matched = false
  if (!wildcards) {
    const needle = oracleCanonical(pattern, locale)
    for (let i = 0; i <= n && !matched; i++) {
      for (let j = i; j <= n && !matched; j++) {
        if (positive === 'eq' && (i !== 0 || j !== n)) continue
        if (positive === 'begins' && i !== 0) continue
        if (positive === 'ends' && j !== n) continue
        matched = slice(i, j) === needle
      }
    }
  } else {
    const tokens = oracleTokens(pattern).map(token => token.kind === 'literal' ? { kind: 'literal' as const, text: oracleCanonical(token.text, locale) } : token)
    if (positive === 'ends' || positive === 'contains') tokens.unshift({ kind: 'star' })
    if (positive === 'begins' || positive === 'contains') tokens.push({ kind: 'star' })
    const memo = new Map<number, boolean>()
    const from = (t: number, g: number): boolean => {
      if (t === tokens.length) return g === n
      const key = t * (n + 1) + g, known = memo.get(key)
      if (known !== undefined) return known
      const token = tokens[t]
      let result = false
      if (token.kind === 'star') result = from(t + 1, g) || (g < n && from(t, g + 1))
      else if (token.kind === 'one') result = g < n && from(t + 1, g + 1)
      else for (let j = g; j <= n && !result; j++) result = slice(g, j) === token.text && from(t + 1, j)
      memo.set(key, result)
      return result
    }
    matched = from(0, 0)
  }
  return negative ? !matched : matched
}
